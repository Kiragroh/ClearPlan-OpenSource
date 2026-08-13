using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;

namespace ClearPlan.Core.Tests
{
    internal static class SyntheticDvhMetricsTests
    {
        private static IList<ReviewDvhPoint> Triangle()
        {
            return new List<ReviewDvhPoint>
            {
                new ReviewDvhPoint(0.0, 100.0),
                new ReviewDvhPoint(10.0, 0.0)
            };
        }

        public static void InterpolatesEndpointsAndMidpoints()
        {
            IList<ReviewDvhPoint> points = Triangle();
            TestAssert.Equal(
                100.0,
                SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(
                    points,
                    0.0));
            TestAssert.Equal(
                50.0,
                SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(
                    points,
                    5.0));
            TestAssert.Equal(
                0.0,
                SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(
                    points,
                    10.0));
            TestAssert.Equal(
                5.0,
                SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(
                    points,
                    50.0));
        }

        public static void IntegratesMeanDose()
        {
            TestAssert.Equal(
                5.0,
                SyntheticDvhMetrics.MeanDoseGy(Triangle()));
        }

        public static void ClassifiesExactThresholdBoundaries()
        {
            TestAssert.Equal(
                ReviewStatusCodes.Pass,
                SyntheticDvhMetrics.EvaluateThresholdStatus(
                    95.0,
                    ">=",
                    95.0,
                    90.0));
            TestAssert.Equal(
                ReviewStatusCodes.Variation,
                SyntheticDvhMetrics.EvaluateThresholdStatus(
                    90.0,
                    ">=",
                    95.0,
                    90.0));
            TestAssert.Equal(
                ReviewStatusCodes.Pass,
                SyntheticDvhMetrics.EvaluateThresholdStatus(
                    20.0,
                    "<=",
                    20.0,
                    25.0));
            TestAssert.Equal(
                ReviewStatusCodes.Variation,
                SyntheticDvhMetrics.EvaluateThresholdStatus(
                    25.0,
                    "<=",
                    20.0,
                    25.0));
        }

        public static void RejectsInvalidCurvesAndRequests()
        {
            TestAssert.Throws<ArgumentException>(
                () => SyntheticDvhMetrics.MeanDoseGy(
                    new List<ReviewDvhPoint>
                    {
                        new ReviewDvhPoint(0.0, 100.0),
                        new ReviewDvhPoint(0.0, 50.0)
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SyntheticDvhMetrics.MeanDoseGy(
                    new List<ReviewDvhPoint>
                    {
                        new ReviewDvhPoint(0.0, 50.0),
                        new ReviewDvhPoint(1.0, 75.0)
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SyntheticDvhMetrics.MeanDoseGy(
                    new List<ReviewDvhPoint>
                    {
                        new ReviewDvhPoint(0.0, 100.0),
                        new ReviewDvhPoint(
                            1.0,
                            double.NaN)
                    }));
            TestAssert.Throws<ArgumentOutOfRangeException>(
                () =>
                    SyntheticDvhMetrics.RelativeVolumePercentAtDoseGy(
                        Triangle(),
                        11.0));
            TestAssert.Throws<ArgumentOutOfRangeException>(
                () =>
                    SyntheticDvhMetrics.DoseGyAtRelativeVolumePercent(
                        Triangle(),
                        101.0));
            TestAssert.Throws<ArgumentException>(
                () =>
                    SyntheticDvhMetrics.EvaluateThresholdStatus(
                        1.0,
                        "==",
                        1.0,
                        1.0));
        }
    }
}
