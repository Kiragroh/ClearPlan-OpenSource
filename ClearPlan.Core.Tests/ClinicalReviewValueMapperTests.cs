using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class ClinicalReviewValueMapperTests
    {
        public static void MapsPlanCheckStatusAndSeverity()
        {
            AssertStatus(
                "3 - OK",
                ReviewStatusCodes.Pass,
                ReviewSeverityCodes.None);
            AssertStatus(
                "2 - Variation",
                ReviewStatusCodes.Variation,
                ReviewSeverityCodes.Warning);
            AssertStatus(
                "1 - Deviation",
                ReviewStatusCodes.Fail,
                ReviewSeverityCodes.Error);
            AssertStatus(
                "0 - Information",
                ReviewStatusCodes.Info,
                ReviewSeverityCodes.Info);
            AssertStatus(
                "unexpected",
                ReviewStatusCodes.NotEvaluated,
                ReviewSeverityCodes.Info);
        }

        public static void MapsPqmStatusAndSeverity()
        {
            AssertPqm(
                "Goal",
                ReviewStatusCodes.Pass,
                ReviewSeverityCodes.None);
            AssertPqm(
                "Variation",
                ReviewStatusCodes.Variation,
                ReviewSeverityCodes.Warning);
            AssertPqm(
                "Not met",
                ReviewStatusCodes.Fail,
                ReviewSeverityCodes.Error);
            AssertPqm(
                "Not evaluated",
                ReviewStatusCodes.NotEvaluated,
                ReviewSeverityCodes.Info);
            AssertPqm(
                null,
                ReviewStatusCodes.NotEvaluated,
                ReviewSeverityCodes.Info);
        }

        public static void CreatesDeterministicUniqueStableIds()
        {
            var used = new HashSet<string>(StringComparer.Ordinal);

            TestAssert.Equal(
                "plan-check",
                ClinicalReviewValueMapper.CreateUniqueStableId(
                    "Plan Check",
                    "check",
                    used));
            TestAssert.Equal(
                "plan-check-2",
                ClinicalReviewValueMapper.CreateUniqueStableId(
                    "Plan Check",
                    "check",
                    used));
            TestAssert.Equal(
                "check",
                ClinicalReviewValueMapper.CreateUniqueStableId(
                    null,
                    "check",
                    used));
            TestAssert.Equal(
                "check-2",
                ClinicalReviewValueMapper.CreateUniqueStableId(
                    @"\\private-server\clinical\constraints.xlsx",
                    "check",
                    used));
        }

        public static void ParsesInvariantAndGermanDecimals()
        {
            double value;
            TestAssert.True(
                ClinicalReviewValueMapper.TryParseClinicalDouble(
                    "<= 12.5 Gy",
                    out value));
            TestAssert.Equal(12.5, value);

            TestAssert.True(
                ClinicalReviewValueMapper.TryParseClinicalDouble(
                    ">= 12,5 Gy",
                    out value));
            TestAssert.Equal(12.5, value);

            TestAssert.False(
                ClinicalReviewValueMapper.TryParseClinicalDouble(
                    "Not evaluated",
                    out value));
        }

        public static void ConvertsCentigrayToGray()
        {
            TestAssert.Equal("%", ClinicalReviewValueMapper.InferUnit("V26Gy[%]"));
            TestAssert.Equal("Gy", ClinicalReviewValueMapper.InferUnit("D2cc[Gy]"));
            TestAssert.Equal(ReviewUnitCodes.CubicCentimeter, ClinicalReviewValueMapper.InferUnit("V26Gy[cc]"));
            TestAssert.Equal(
                2.5,
                ClinicalReviewValueMapper.ConvertDoseToGray(250.0, "cGy"));
            TestAssert.Equal(
                2.5,
                ClinicalReviewValueMapper.ConvertDoseToGray(2.5, "Gy"));
            TestAssert.Throws<ArgumentException>(
                () => ClinicalReviewValueMapper.ConvertDoseToGray(2.5, "mystery"));
        }

        public static void MapsFieldConformance()
        {
            TestAssert.Equal(
                ReviewStatusCodes.Pass,
                ClinicalReviewValueMapper.MapFieldStatus("7GA01", "7GA01"));
            TestAssert.Equal(
                ReviewStatusCodes.Variation,
                ClinicalReviewValueMapper.MapFieldStatus("7GA01", "7GA02"));
            TestAssert.Equal(
                ReviewStatusCodes.NotEvaluated,
                ClinicalReviewValueMapper.MapFieldStatus(null, "7GA02"));
        }

        public static void MissingPlanDoseDoesNotAbortOverview()
        {
            var read = typeof(ClinicalReviewValueMapper).GetMethod("ReadOptionalDoseInGray");
            TestAssert.NotNull(read, "Optional native plan dose must not abort the whole overview.");
            Func<Func<double>, double?> capture = provider => (double?)read.Invoke(null, new object[] { provider });
            foreach (double missing in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity, -1.0 })
            {
                TestAssert.Equal<double?>(null, capture(() => missing));
                TestAssert.Equal<double?>(null, capture(() => ClinicalReviewValueMapper.ConvertDoseToGray(missing, "Gy")));
            }
            TestAssert.Equal<double?>(null, capture(() => { throw new InvalidOperationException("Native value unavailable"); }));
            TestAssert.Equal<double?>(null, capture(() => ClinicalReviewValueMapper.ConvertDoseToGray(100, "%")));
            TestAssert.Equal<double?>(0.0, capture(() => 0.0));
            TestAssert.Equal<double?>(2.5, capture(() => ClinicalReviewValueMapper.ConvertDoseToGray(250, "cGy")));
            TestAssert.Equal<double?>(70.0, capture(() => ClinicalReviewValueMapper.ConvertDoseToGray(70, "Gy")));
            // Missing total dose does not erase an independently readable prescription.
            var plan = new ReviewPlanRow { DosePerFractionGy = capture(() => 2.0),
                TotalDoseGy = capture(() => ClinicalReviewValueMapper.ConvertDoseToGray(double.NaN, "Gy")), FractionCount = 33 };
            TestAssert.Equal<double?>(2, plan.DosePerFractionGy);
            TestAssert.Equal<double?>(null, plan.TotalDoseGy);
            TestAssert.Equal<int?>(33, plan.FractionCount);
            // The strict converter still rejects invalid values for actual calculations.
            TestAssert.Throws<ArgumentOutOfRangeException>(() => ClinicalReviewValueMapper.ConvertDoseToGray(double.NaN, "Gy"));
            var root = new System.IO.DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (root != null && !System.IO.File.Exists(System.IO.Path.Combine(root.FullName, "ClearPlan.sln"))) root = root.Parent;
            TestAssert.NotNull(root);
            string adapter = System.IO.File.ReadAllText(System.IO.Path.Combine(root.FullName, "ClearPlan.Script", "Review", "EsapiReviewSnapshotBuilder.cs"));
            TestAssert.True(adapter.Contains("GetDvhDoseOrNull(() => planSetup.DosePerFraction)"));
            TestAssert.True(adapter.Contains("GetDvhDoseOrNull(() => planSetup.TotalDose)"));
            // Private headless-runner integration is tested in the local deployment,
            // not shipped in the vendor-free public source/test package.
        }

        public static void SanitizesPathsAndDicomUids()
        {
            TestAssert.Equal(
                "Constraint source",
                ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    @"\\private-server\clinical\constraints.xlsx",
                    "Constraint source"));
            TestAssert.Equal(
                "Constraint source",
                ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    @"C:\clinical\constraints.xlsx",
                    "Constraint source"));
            TestAssert.Equal(
                "Clinical item",
                ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    "1.2.840.10008.5.1.4.1.1.481.5",
                    "Clinical item"));
            TestAssert.Equal(
                "Plan C2",
                ClinicalReviewValueMapper.SanitizeClinicalLabel(
                    "Plan C2",
                    "Clinical item"));
        }

        private static void AssertStatus(
            string source,
            string expectedStatus,
            string expectedSeverity)
        {
            ReviewStatusAndSeverity result =
                ClinicalReviewValueMapper.MapPlanCheckStatus(source);
            TestAssert.Equal(expectedStatus, result.Status);
            TestAssert.Equal(expectedSeverity, result.Severity);
        }

        private static void AssertPqm(
            string source,
            string expectedStatus,
            string expectedSeverity)
        {
            ReviewStatusAndSeverity result =
                ClinicalReviewValueMapper.MapPqmStatus(source);
            TestAssert.Equal(expectedStatus, result.Status);
            TestAssert.Equal(expectedSeverity, result.Severity);
        }
    }
}
