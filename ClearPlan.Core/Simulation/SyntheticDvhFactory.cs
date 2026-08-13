using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Simulation
{
    public static class SyntheticDvhFactory
    {
        public static List<ReviewDvhPoint> CreateLogisticCumulativePercent(
            double maximumDoseGy,
            double doseStepGy,
            double midpointDoseGy,
            double slopeGy)
        {
            ValidateGrid(maximumDoseGy, doseStepGy);
            if (slopeGy <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                    "slopeGy",
                    "The logistic slope in Gy must be positive.");
            }

            return CreateGrid(
                maximumDoseGy,
                doseStepGy,
                delegate(double doseGy)
                {
                    return 100.0 /
                           (1.0 + Math.Exp(
                               (doseGy - midpointDoseGy) / slopeGy));
                });
        }

        public static List<ReviewDvhPoint> CreateExponentialCumulativePercent(
            double maximumDoseGy,
            double doseStepGy,
            double scaleDoseGy,
            double shape)
        {
            ValidateGrid(maximumDoseGy, doseStepGy);
            if (scaleDoseGy <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                    "scaleDoseGy",
                    "The exponential scale in Gy must be positive.");
            }

            if (shape <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                    "shape",
                    "The exponential shape must be positive.");
            }

            return CreateGrid(
                maximumDoseGy,
                doseStepGy,
                delegate(double doseGy)
                {
                    return 100.0 * Math.Exp(
                        -Math.Pow(doseGy / scaleDoseGy, shape));
                });
        }

        private static List<ReviewDvhPoint> CreateGrid(
            double maximumDoseGy,
            double doseStepGy,
            Func<double, double> relativeVolumePercent)
        {
            int intervalCount = (int)Math.Round(
                maximumDoseGy / doseStepGy,
                MidpointRounding.AwayFromZero);
            var points = new List<ReviewDvhPoint>(intervalCount + 1);
            double previousVolumePercent = 100.0;
            for (int index = 0; index <= intervalCount; index++)
            {
                double doseGy = index == intervalCount
                    ? maximumDoseGy
                    : index * doseStepGy;
                double volumePercent = index == 0
                    ? 100.0
                    : Math.Max(
                        0.0,
                        Math.Min(
                            previousVolumePercent,
                            relativeVolumePercent(doseGy)));
                doseGy = Math.Round(doseGy, 3);
                volumePercent = Math.Round(volumePercent, 3);
                points.Add(new ReviewDvhPoint(doseGy, volumePercent));
                previousVolumePercent = volumePercent;
            }

            return points;
        }

        private static void ValidateGrid(
            double maximumDoseGy,
            double doseStepGy)
        {
            if (maximumDoseGy <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                    "maximumDoseGy",
                    "The maximum dose in Gy must be positive.");
            }

            if (doseStepGy <= 0.0 || doseStepGy > maximumDoseGy)
            {
                throw new ArgumentOutOfRangeException(
                    "doseStepGy",
                    "The dose step in Gy must be positive and no larger than the maximum dose.");
            }

            double intervalCount = maximumDoseGy / doseStepGy;
            if (Math.Abs(intervalCount - Math.Round(intervalCount)) > 1e-9)
            {
                throw new ArgumentException(
                    "The maximum dose must be an integer multiple of the dose step.");
            }
        }
    }
}
