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
