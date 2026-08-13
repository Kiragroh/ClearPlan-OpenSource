using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class PqmDefaultReviewPolicyTests
    {
        public static void KeepsValidRowsVisible()
        {
            TestAssert.False(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Trachea",
                    "0.0 Gy",
                    "Goal"));
            TestAssert.False(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Heart",
                    "0.00 %",
                    "Goal"));
            TestAssert.False(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Control",
                    "12.0 Gy",
                    "Variation"));
        }

        public static void IgnoresDelimitedPartialHelpers()
        {
            TestAssert.True(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Parotid_TR",
                    "12.0 Gy",
                    "Goal"));
            TestAssert.True(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Lunge teil",
                    "8.0 Gy",
                    "Goal"));
            TestAssert.False(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Trachea",
                    "8.0 Gy",
                    "Goal"));
            TestAssert.False(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "PTV_TR",
                    "8.0 Gy",
                    "Goal"));
        }

        public static void UsesExplicitUnavailableStatusForNonTargets()
        {
            TestAssert.True(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Kidney_L",
                    "Not evaluated",
                    "Not evaluated"));
            TestAssert.True(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "Cochlea_R",
                    "Volume too small",
                    "Not evaluated"));
            TestAssert.False(
                PqmDefaultReviewPolicy.ShouldIgnore(
                    "PTV_60",
                    "Not evaluated",
                    "Not evaluated"));
        }
    }
}
