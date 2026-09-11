using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ClearPlan.Core.Review;
using ClearPlan.Reporting;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewDvhStatisticsTests
    {
        public static void RunAll()
        {
            InterpolatesDosePercentiles();
            UsesRightEdgeOfInteriorPlateaus();
            PrefersNativeStatistics();
            EstimatesSyntheticStatisticsFromFullCoverage();
            DoesNotInventMissingClinicalStatistics();
            RejectsInvalidCurves();
            LeavesUncoveredPercentilesUnavailable();
            RejectsInvalidNativeStatistics();
            MapsNativeMetadataAndSyntheticEstimatesSeparately();
            NativeOarD98OverridesCoarseCurveInterpolation();
            NativeCaptureUsesUnitAwareProviderStatistics();
        }

        public static void InterpolatesDosePercentiles()
        {
            object result = Calculate(Curve(0, 100, 10, 0));
            Near(0.2, Number(result, "D98Gy"));
            Near(5, Number(result, "MedianDoseGy"));
            Near(9.8, Number(result, "D2Gy"));
        }

        public static void UsesRightEdgeOfInteriorPlateaus()
        {
            var points = Curve(0, 100, 10, 98, 20, 98, 30, 50, 40, 50, 50, 2, 60, 2, 70, 0);
            object result = Calculate(points);
            Near(20, Number(result, "D98Gy"));
            Near(40, Number(result, "MedianDoseGy"));
            Near(60, Number(result, "D2Gy"));
            TestAssert.Equal(8, points.Count, "Statistics must not remove plateau samples.");
        }

        public static void PrefersNativeStatistics()
        {
            object result = Calculate(Curve(0, 100, 10, 0), 0.1, 4.9, 9.9, true);
            Near(0.1, Number(result, "MinimumDoseGy"));
            Near(4.9, Number(result, "MeanDoseGy"));
            Near(9.9, Number(result, "MaximumDoseGy"));
            TestAssert.Equal(false, Property(result, "MinimumDoseEstimated"));
            TestAssert.Equal(false, Property(result, "MeanDoseEstimated"));
            TestAssert.Equal(false, Property(result, "MaximumDoseEstimated"));
        }

        public static void EstimatesSyntheticStatisticsFromFullCoverage()
        {
            object result = Calculate(Curve(4, 100, 10, 100, 20, 0, 30, 0), null, null, null, true);
            Near(10, Number(result, "MinimumDoseGy"));
            Near(15, Number(result, "MeanDoseGy"));
            Near(20, Number(result, "MaximumDoseGy"));
            TestAssert.Equal(true, Property(result, "MinimumDoseEstimated"));
            TestAssert.Equal(true, Property(result, "MeanDoseEstimated"));
            TestAssert.Equal(true, Property(result, "MaximumDoseEstimated"));
        }

        public static void DoesNotInventMissingClinicalStatistics()
        {
            object result = Calculate(Curve(0, 100, 10, 0));
            TestAssert.Equal(null, Number(result, "MinimumDoseGy"));
            TestAssert.Equal(null, Number(result, "MeanDoseGy"));
            TestAssert.Equal(null, Number(result, "MaximumDoseGy"));
            Near(5, Number(result, "MedianDoseGy"));
        }

        public static void RejectsInvalidCurves()
        {
            var invalid = new List<List<ReviewDvhPoint>>
            {
                null, new List<ReviewDvhPoint>(), Curve(0, 100),
                Curve(0, 100, 10, 50, 9, 0), Curve(0, 100, 0, 0),
                Curve(0, 100, 1, 40, 2, 50, 10, 0),
                Curve(-1, 100, 10, 0), Curve(0, 101, 10, 0),
                Curve(0, 100, 10, -1), Curve(0, 100, double.NaN, 0),
                Curve(0, 100, 10, double.PositiveInfinity),
                new List<ReviewDvhPoint> { new ReviewDvhPoint(0, 100), null }
            };
            foreach (var points in invalid)
            {
                object result = Calculate(points, null, null, null, true);
                foreach (string metric in new[] { "MinimumDoseGy", "D98Gy", "MeanDoseGy", "MedianDoseGy", "MaximumDoseGy", "D2Gy" })
                    TestAssert.Equal(null, Number(result, metric), "Invalid curve must not yield " + metric + ".");
            }
        }

        public static void LeavesUncoveredPercentilesUnavailable()
        {
            object result = Calculate(Curve(1, 95, 9, 5), null, null, null, true);
            TestAssert.Equal(null, Number(result, "D98Gy"));
            Near(5, Number(result, "MedianDoseGy"));
            TestAssert.Equal(null, Number(result, "D2Gy"));
            TestAssert.Equal(null, Number(result, "MinimumDoseGy"));
            TestAssert.Equal(null, Number(result, "MeanDoseGy"));
            TestAssert.Equal(null, Number(result, "MaximumDoseGy"));
            foreach (var points in new[] { Curve(0, 100, 10, 100), Curve(0, 0, 10, 0) })
            {
                result = Calculate(points, null, null, null, true);
                TestAssert.Equal(null, Number(result, "MinimumDoseGy"));
                TestAssert.Equal(null, Number(result, "MeanDoseGy"));
                TestAssert.Equal(null, Number(result, "MaximumDoseGy"));
            }
        }

        public static void RejectsInvalidNativeStatistics()
        {
            object result = Calculate(Curve(0, 100, 10, 0), double.NaN, -1, double.PositiveInfinity);
            TestAssert.Equal(null, Number(result, "MinimumDoseGy"));
            TestAssert.Equal(null, Number(result, "MeanDoseGy"));
            TestAssert.Equal(null, Number(result, "MaximumDoseGy"));
            result = Calculate(null, 0, 5, 10);
            Near(0, Number(result, "MinimumDoseGy"));
            Near(5, Number(result, "MeanDoseGy"));
            Near(10, Number(result, "MaximumDoseGy"));
            TestAssert.Equal(null, Number(result, "MedianDoseGy"));
            result = Calculate(Curve(0, 100, 10, 0), 8, 5, 4);
            TestAssert.Equal(null, Number(result, "MinimumDoseGy"));
            TestAssert.Equal(null, Number(result, "MeanDoseGy"));
            TestAssert.Equal(null, Number(result, "MaximumDoseGy"));
        }

        public static void MapsNativeMetadataAndSyntheticEstimatesSeparately()
        {
            var snapshot = new ReviewSnapshot { Synthetic = false };
            var series = new ReviewDvhSeries { Points = Curve(0, 100, 10, 0) };
            snapshot.DvhSeries.Add(series);
            object missing = Property(new ReviewSnapshotReportMapper().Map(snapshot).DvhSeries[0], "Statistics");
            TestAssert.Equal(null, Number(missing, "MeanDoseGy"));
            Set(series, "MinimumDoseGy", 0.1);
            Set(series, "MeanDoseGy", 4.9);
            Set(series, "MaximumDoseGy", 9.9);
            object native = Property(new ReviewSnapshotReportMapper().Map(snapshot).DvhSeries[0], "Statistics");
            Near(0.1, Number(native, "MinimumDoseGy"));
            Near(4.9, Number(native, "MeanDoseGy"));
            Near(9.9, Number(native, "MaximumDoseGy"));
            Near(5, Number(native, "MedianDoseGy"));
            Set(series, "MeanDoseGy", 4.8);
            Near(4.9, Number(native, "MeanDoseGy"));
            Set(series, "MinimumDoseGy", null);
            Set(series, "MeanDoseGy", null);
            Set(series, "MaximumDoseGy", null);
            snapshot.Synthetic = true;
            object synthetic = Property(new ReviewSnapshotReportMapper().Map(snapshot).DvhSeries[0], "Statistics");
            Near(5, Number(synthetic, "MeanDoseGy"));
            TestAssert.Equal(true, Property(synthetic, "MeanDoseEstimated"));
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(snapshot);
            TestAssert.True(json.Contains("meanDoseGy"), "Native metadata must round-trip in detached JSON.");
        }

        public static void NativeOarD98OverridesCoarseCurveInterpolation()
        {
            var snapshot = new ReviewSnapshot { Synthetic = false };
            // Synthetic low-dose fixture: coarse interpolation would yield D98 = 0.04 Gy,
            // below native Dmin. Retain the direct provider value without clamping.
            snapshot.DvhSeries.Add(new ReviewDvhSeries {
                StructureId = "Synthetic_OAR", Role = ReviewDvhRoleCodes.OrganAtRisk,
                Points = Curve(0, 100, 2, 0), MinimumDoseGy = 0.12, MeanDoseGy = 0.8,
                MaximumDoseGy = 1.9, D98DoseGy = 0.18
            });
            var statistics = new ReviewSnapshotReportMapper().Map(snapshot).DvhSeries[0].Statistics;
            Near(0.18, statistics.D98Gy);
            Near(0.12, statistics.MinimumDoseGy);
            Near(1.0, statistics.MedianDoseGy);
            Near(1.96, statistics.D2Gy);
            snapshot.DvhSeries[0].D98DoseGy = 0;
            Near(0, new ReviewSnapshotReportMapper().Map(snapshot).DvhSeries[0].Statistics.D98Gy);
        }

        public static void NativeCaptureUsesUnitAwareProviderStatistics()
        {
            DirectoryInfo directory = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (directory != null && !File.Exists(Path.Combine(directory.FullName, "ClearPlan.sln"))) directory = directory.Parent;
            if (directory == null) directory = new DirectoryInfo(Environment.CurrentDirectory);
            string code = File.ReadAllText(Path.Combine(directory.FullName, "ClearPlan.Script", "Review", "EsapiReviewSnapshotBuilder.cs"));
            foreach (string expression in new[] { "GetDvhDoseOrNull(() => data.MinDose)", "GetDvhDoseOrNull(() => data.MeanDose)", "GetDvhDoseOrNull(() => data.MaxDose)", "ConvertDoseToGray(readDose())" })
                TestAssert.True(code.Contains(expression), "Native DVH capture must use " + expression);
            TestAssert.True(code.Contains("double? nativeD98 = EsapiTargetReviewBuilder.ReadDoseAtVolume("),
                "Native D98 capture must include OARs, not be restricted to target structures.");
        }

        private static object Calculate(IList<ReviewDvhPoint> points, double? min = null, double? mean = null, double? max = null, bool estimates = false)
        {
            Type calculator = typeof(ReviewDvhPoint).Assembly.GetType("ClearPlan.Core.Review.ReviewDvhStatisticsCalculator");
            TestAssert.NotNull(calculator, "DVH statistic calculator has not been implemented.");
            return calculator.GetMethod("Calculate").Invoke(null, new object[] { points, min, mean, max, estimates });
        }

        private static object Property(object result, string name)
        {
            TestAssert.NotNull(result);
            PropertyInfo property = result.GetType().GetProperty(name);
            TestAssert.NotNull(property, "Missing DVH property: " + name);
            return property.GetValue(result, null);
        }

        private static void Set(object target, string name, object value)
        {
            PropertyInfo property = target.GetType().GetProperty(name);
            TestAssert.NotNull(property, "Missing native DVH property: " + name);
            property.SetValue(target, value, null);
        }

        private static double? Number(object result, string name) { return (double?)Property(result, name); }
        private static void Near(double expected, double? actual)
        {
            TestAssert.True(actual.HasValue && Math.Abs(actual.Value - expected) < 1e-10,
                string.Format("Expected DVH dose {0}, got {1}.", expected, actual));
        }

        private static List<ReviewDvhPoint> Curve(params double[] values)
        {
            var result = new List<ReviewDvhPoint>();
            for (int i = 0; i < values.Length; i += 2) result.Add(new ReviewDvhPoint(values[i], values[i + 1]));
            return result;
        }
    }
}
