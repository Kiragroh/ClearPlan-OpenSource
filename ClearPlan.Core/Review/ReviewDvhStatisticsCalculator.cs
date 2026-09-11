using System;
using System.Collections.Generic;

namespace ClearPlan.Core.Review
{
    /// <summary>Detached absolute-dose summary. Null denotes unavailable, never zero.</summary>
    public sealed class ReviewDvhStatistics
    {
        public double? MinimumDoseGy { get; set; }
        public double? D98Gy { get; set; }
        public double? MeanDoseGy { get; set; }
        public double? MedianDoseGy { get; set; }
        public double? MaximumDoseGy { get; set; }
        public double? D2Gy { get; set; }
        public bool MinimumDoseEstimated { get; set; }
        public bool MeanDoseEstimated { get; set; }
        public bool MaximumDoseEstimated { get; set; }
    }

    /// <summary>
    /// Summarizes a cumulative DVH with dose in Gy and volume in percent.
    /// D98/D50/D2 use linear interpolation and the right edge of an exact
    /// volume plateau (the greatest sampled/interpolated dose covering that volume).
    /// Native extrema/mean take precedence. Curve estimates are opt-in for simulation.
    /// </summary>
    public static class ReviewDvhStatisticsCalculator
    {
        public static ReviewDvhStatistics Calculate(
            IList<ReviewDvhPoint> points,
            double? minimumDoseGy = null,
            double? meanDoseGy = null,
            double? maximumDoseGy = null,
            bool allowCurveEstimates = false)
        {
            var result = new ReviewDvhStatistics
            {
                MinimumDoseGy = ValidDoseOrNull(minimumDoseGy),
                MeanDoseGy = ValidDoseOrNull(meanDoseGy),
                MaximumDoseGy = ValidDoseOrNull(maximumDoseGy)
            };
            if (result.MinimumDoseGy > result.MeanDoseGy ||
                result.MeanDoseGy > result.MaximumDoseGy ||
                result.MinimumDoseGy > result.MaximumDoseGy)
            {
                // Contradictory native metadata must not look like valid statistics.
                result.MinimumDoseGy = result.MeanDoseGy = result.MaximumDoseGy = null;
            }

            if (!IsValidCurve(points)) return result;

            result.D98Gy = DoseAtVolume(points, 98.0);
            result.MedianDoseGy = DoseAtVolume(points, 50.0);
            result.D2Gy = DoseAtVolume(points, 2.0);

            // Do not extrapolate a missing low-dose or high-dose tail. A constant
            // 100% or 0% curve is not sufficient evidence for any of these estimates.
            if (!allowCurveEstimates || points[0].VolumePercent != 100.0 ||
                points[points.Count - 1].VolumePercent != 0.0) return result;

            if (!result.MinimumDoseGy.HasValue)
            {
                result.MinimumDoseGy = DoseAtVolume(points, 100.0);
                result.MinimumDoseEstimated = result.MinimumDoseGy.HasValue;
            }
            if (!result.MeanDoseGy.HasValue)
            {
                // A curve starting above zero at 100% implies a 100% interval
                // from zero to the first dose; include it in the survival integral.
                double mean = points[0].DoseGy;
                for (int index = 1; index < points.Count; index++)
                {
                    ReviewDvhPoint previous = points[index - 1];
                    ReviewDvhPoint current = points[index];
                    mean += (current.DoseGy - previous.DoseGy) *
                        ((previous.VolumePercent + current.VolumePercent) / 200.0);
                }
                result.MeanDoseGy = ValidDoseOrNull(mean);
                result.MeanDoseEstimated = result.MeanDoseGy.HasValue;
            }
            if (!result.MaximumDoseGy.HasValue)
            {
                // The first zero-volume bin bounds the high-dose tail, not the
                // last plotted dose when the curve contains a trailing zero plateau.
                for (int index = 1; index < points.Count; index++)
                {
                    if (points[index].VolumePercent != 0.0) continue;
                    result.MaximumDoseGy = points[index].DoseGy;
                    result.MaximumDoseEstimated = true;
                    break;
                }
            }
            return result;
        }

        private static double? DoseAtVolume(IList<ReviewDvhPoint> points, double volumePercent)
        {
            if (points[0].VolumePercent < volumePercent ||
                points[points.Count - 1].VolumePercent > volumePercent) return null;

            // Strictly less skips every sample on an exact-volume plateau.
            for (int index = 1; index < points.Count; index++)
            {
                ReviewDvhPoint current = points[index];
                if (current.VolumePercent >= volumePercent) continue;
                ReviewDvhPoint previous = points[index - 1];
                double fraction = (previous.VolumePercent - volumePercent) /
                    (previous.VolumePercent - current.VolumePercent);
                return ValidDoseOrNull(previous.DoseGy +
                    fraction * (current.DoseGy - previous.DoseGy));
            }
            // The target volume at a right-censored plateau has no observed
            // descending bracket. Its true right edge is not in this curve.
            return null;
        }

        private static bool IsValidCurve(IList<ReviewDvhPoint> points)
        {
            if (points == null || points.Count < 2) return false;
            ReviewDvhPoint previous = null;
            foreach (ReviewDvhPoint point in points)
            {
                if (point == null || !ValidDoseOrNull(point.DoseGy).HasValue ||
                    double.IsNaN(point.VolumePercent) || double.IsInfinity(point.VolumePercent) ||
                    point.VolumePercent < 0.0 || point.VolumePercent > 100.0 ||
                    (previous != null && (point.DoseGy <= previous.DoseGy ||
                        point.VolumePercent > previous.VolumePercent))) return false;
                previous = point;
            }
            return true;
        }

        private static double? ValidDoseOrNull(double? value)
        {
            return value.HasValue && !double.IsNaN(value.Value) &&
                !double.IsInfinity(value.Value) && value.Value >= 0.0 ? value : null;
        }
    }
}
