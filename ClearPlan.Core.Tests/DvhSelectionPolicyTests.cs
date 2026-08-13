using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class DvhSelectionPolicyTests
    {
        public static void SelectsExactMappedStructures()
        {
            var requested = new[] { "Heart", "PTV_20" };

            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Heart", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("ptv_20", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("Heart_PRV", requested));
        }

        public static void ExcludesExternalAndBodyContours()
        {
            var requested = new[] { "External", "BODY", "Körper", "Lung_L" };

            TestAssert.False(DvhSelectionPolicy.ShouldSelect("External", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("BODY", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("Körper", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("Lung_L", requested));
        }

        public static void ExcludesPartialHelpersButKeepsTargets()
        {
            var requested = new[] { "Lunge Teil", "Heart_TL", "PTV_TL" };

            TestAssert.False(DvhSelectionPolicy.ShouldSelect("Lunge Teil", requested));
            TestAssert.False(DvhSelectionPolicy.ShouldSelect("Heart_TL", requested));
            TestAssert.True(DvhSelectionPolicy.ShouldSelect("PTV_TL", requested));
        }
    }
}
