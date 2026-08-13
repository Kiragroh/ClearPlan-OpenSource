using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Simulation
{
    public static class SyntheticDvhMetrics
    {
        public static double RelativeVolumePercentAtDoseGy(
            IList<ReviewDvhPoint> points,
            double doseGy)
        {
            ValidatePoints(points);
            if (doseGy < points[0].DoseGy ||
                doseGy > points[points.Count - 1].DoseGy)
            {
                throw new ArgumentOutOfRangeException(
                    "doseGy",
                    "The requested dose must lie inside the sampled DVH.");
            }

            for (int index = 1; index < points.Count; index++)
            {
                ReviewDvhPoint previous = points[index - 1];
                ReviewDvhPoint current = points[index];
                if (doseGy <= current.DoseGy)
                {
                    double fraction =
                        (doseGy - previous.DoseGy) /
                        (current.DoseGy - previous.DoseGy);
                    return previous.VolumePercent +
                           fraction *
                           (current.VolumePercent - previous.VolumePercent);
                }
            }

            return points[points.Count - 1].VolumePercent;
        }

        public static double DoseGyAtRelativeVolumePercent(
            IList<ReviewDvhPoint> points,
            double volumePercent)
        {
            ValidatePoints(points);
            if (volumePercent > points[0].VolumePercent ||
                volumePercent < points[points.Count - 1].VolumePercent)
            {
                throw new ArgumentOutOfRangeException(
                    "volumePercent",
                    "The requested relative volume must lie inside the sampled DVH.");
            }

            for (int index = 1; index < points.Count; index++)
            {
                ReviewDvhPoint previous = points[index - 1];
                ReviewDvhPoint current = points[index];
                if (volumePercent >= current.VolumePercent)
                {
                    double volumeDifference =
                        previous.VolumePercent - current.VolumePercent;
                    if (volumeDifference <= 0.0)
                    {
                        return current.DoseGy;
                    }

                    double fraction =
                        (previous.VolumePercent - volumePercent) /
                        volumeDifference;
                    return previous.DoseGy +
                           fraction *
                           (current.DoseGy - previous.DoseGy);
                }
            }

            return points[points.Count - 1].DoseGy;
        }

        public static double MeanDoseGy(IList<ReviewDvhPoint> points)
        {
            ValidatePoints(points);
            double meanDoseGy = 0.0;
            for (int index = 1; index < points.Count; index++)
            {
                ReviewDvhPoint previous = points[index - 1];
                ReviewDvhPoint current = points[index];
                meanDoseGy +=
                    (current.DoseGy - previous.DoseGy) *
                    (previous.VolumePercent + current.VolumePercent) /
                    200.0;
            }

            return meanDoseGy;
        }

        public static string EvaluateThresholdStatus(
            double achieved,
            string comparator,
            double goal,
            double variation)
        {
            if (comparator == ">=")
            {
                if (variation > goal)
                {
                    throw new ArgumentException(
                        "For >= objectives the variation threshold must not exceed the goal.");
                }

                if (achieved >= goal)
                {
                    return ReviewStatusCodes.Pass;
                }

                return achieved >= variation
                    ? ReviewStatusCodes.Variation
                    : ReviewStatusCodes.Fail;
            }

            if (comparator == "<=")
            {
                if (variation < goal)
                {
                    throw new ArgumentException(
                        "For <= objectives the variation threshold must not be below the goal.");
                }

                if (achieved <= goal)
                {
                    return ReviewStatusCodes.Pass;
                }

                return achieved <= variation
                    ? ReviewStatusCodes.Variation
                    : ReviewStatusCodes.Fail;
            }

            throw new ArgumentException(
                "Only >= and <= synthetic PQM comparators are supported.",
                "comparator");
        }

        private static void ValidatePoints(IList<ReviewDvhPoint> points)
        {
            if (points == null || points.Count < 2)
            {
                throw new ArgumentException(
                    "At least two sampled DVH points are required.",
                    "points");
            }

            ReviewDvhPoint previous = null;
            for (int index = 0; index < points.Count; index++)
            {
                ReviewDvhPoint current = points[index];
                if (current == null ||
                    double.IsNaN(current.DoseGy) ||
                    double.IsInfinity(current.DoseGy) ||
                    current.DoseGy < 0.0 ||
                    double.IsNaN(current.VolumePercent) ||
                    double.IsInfinity(current.VolumePercent) ||
                    current.VolumePercent < 0.0 ||
                    current.VolumePercent > 100.0)
                {
                    throw new ArgumentException(
                        "DVH points must contain finite non-negative dose and relative volume between 0 and 100 percent.",
                        "points");
                }

                if (previous != null &&
                    (current.DoseGy <= previous.DoseGy ||
                     current.VolumePercent >
                     previous.VolumePercent))
                {
                    throw new ArgumentException(
                        "DVH dose must increase strictly and cumulative volume must not increase.",
                        "points");
                }

                previous = current;
            }
        }
    }
}
