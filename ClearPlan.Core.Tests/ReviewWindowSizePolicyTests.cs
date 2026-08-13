using System;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewWindowSizePolicyTests
    {
        public static void MatchesTheResponsiveWorkAreaMatrix()
        {
            AssertSize(1180.0, 720.0, 1180.0, 720.0);
            AssertSize(1366.0, 768.0, 1256.72, 720.0);
            AssertSize(1600.0, 900.0, 1472.0, 792.0);
            AssertSize(1920.0, 1080.0, 1680.0, 950.4);
            AssertSize(2560.0, 1440.0, 1680.0, 1267.2);
        }

        public static void ExposesTheWindowSizingContract()
        {
            TestAssert.Equal(
                1180.0,
                ReviewWindowSizePolicy.MinimumWidth);
            TestAssert.Equal(
                720.0,
                ReviewWindowSizePolicy.MinimumHeight);
            TestAssert.Equal(
                1680.0,
                ReviewWindowSizePolicy.MaximumInitialWidth);
            TestAssert.Equal(
                0.92,
                ReviewWindowSizePolicy.WidthFraction);
            TestAssert.Equal(
                0.88,
                ReviewWindowSizePolicy.HeightFraction);
        }

        public static void RejectsInvalidWorkAreas()
        {
            foreach (double invalid in new[]
            {
                double.NaN,
                double.PositiveInfinity,
                double.NegativeInfinity,
                0.0,
                -1.0
            })
            {
                TestAssert.Throws<ArgumentOutOfRangeException>(
                    () => ReviewWindowSizePolicy.Calculate(
                        invalid,
                        900.0));
                TestAssert.Throws<ArgumentOutOfRangeException>(
                    () => ReviewWindowSizePolicy.Calculate(
                        1600.0,
                        invalid));
            }
        }

        private static void AssertSize(
            double availableWidth,
            double availableHeight,
            double expectedWidth,
            double expectedHeight)
        {
            ReviewWindowSize actual =
                ReviewWindowSizePolicy.Calculate(
                    availableWidth,
                    availableHeight);

            TestAssert.True(
                Math.Abs(expectedWidth - actual.Width) < 0.000001,
                "Unexpected responsive review width.");
            TestAssert.True(
                Math.Abs(expectedHeight - actual.Height) < 0.000001,
                "Unexpected responsive review height.");
        }
    }
}
