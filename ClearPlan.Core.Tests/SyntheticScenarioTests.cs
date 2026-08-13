using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;

namespace ClearPlan.Core.Tests
{
    internal static class SyntheticScenarioTests
    {
        private const string SyntheticWatermark =
            "SYNTHETIC DEMONSTRATION — NOT FOR CLINICAL USE";

        private static readonly ExpectedScenario[] ExpectedScenarios =
        {
            new ExpectedScenario(
                "baseline-pass", 1101, 2, 6, 8, 5, 4, 5,
                6, 0, 0,
                8, 0, 0,
                5, 0, 5, 0,
                4, 0,
                2, 0),
            new ExpectedScenario(
                "target-underdose", 1102, 2, 6, 8, 5, 4, 5,
                4, 1, 1,
                8, 0, 0,
                5, 0, 5, 0,
                4, 0,
                2, 0),
            new ExpectedScenario(
                "oar-overdose", 1103, 2, 6, 8, 5, 4, 5,
                4, 1, 1,
                8, 0, 0,
                5, 0, 5, 0,
                4, 0,
                2, 0),
            new ExpectedScenario(
                "metadata-plancheck", 1104, 2, 6, 8, 5, 4, 5,
                6, 0, 0,
                3, 1, 4,
                5, 0, 5, 0,
                4, 0,
                2, 0),
            new ExpectedScenario(
                "field-and-mapping", 1105, 2, 6, 8, 5, 4, 5,
                6, 0, 0,
                8, 0, 0,
                5, 0, 0, 5,
                3, 1,
                2, 0),
            new ExpectedScenario(
                "optional-path-fallback", 1106, 3, 6, 8, 5, 4, 5,
                6, 0, 0,
                8, 0, 0,
                5, 0, 5, 0,
                4, 0,
                2, 1),
            new ExpectedScenario(
                "mixed-review", 1107, 3, 6, 8, 5, 4, 5,
                2, 2, 2,
                3, 1, 4,
                5, 0, 0, 5,
                3, 1,
                2, 1)
        };

        private static readonly string[] RequiredStructures =
        {
            "PTV_60",
            "SpinalCord",
            "Parotid_L",
            "Parotid_R",
            "External"
        };

        public static void DeclaresExactCatalogAndExpectations()
        {
            IList<string> actualIds = SyntheticScenarioFactory.ScenarioIds;
            TestAssert.Equal(ExpectedScenarios.Length, actualIds.Count);

            for (int index = 0; index < ExpectedScenarios.Length; index++)
            {
                ExpectedScenario expected = ExpectedScenarios[index];
                TestAssert.Equal(expected.ScenarioId, actualIds[index]);

                SyntheticScenarioExpectation actual =
                    SyntheticScenarioFactory.GetExpectation(expected.ScenarioId);
                TestAssert.NotNull(actual);
                TestAssert.Equal(expected.Seed, actual.Seed);
                TestAssert.Equal(expected.SourceCount, actual.SourceCount);
                TestAssert.Equal(expected.PqmCount, actual.PqmRowCount);
                TestAssert.Equal(
                    expected.PlanCheckCount,
                    actual.PlanCheckRowCount);
                TestAssert.Equal(expected.FieldCount, actual.FieldRowCount);
                TestAssert.Equal(
                    expected.MappingCount,
                    actual.StructureMappingCount);
                TestAssert.Equal(
                    expected.DvhSeriesCount,
                    actual.DvhSeriesCount);
                TestAssert.Equal(
                    string.Join("|", RequiredStructures),
                    string.Join("|", actual.RequiredStructureIds));
            }
        }

        public static void MatchesExactRowsStatusesAndStructures()
        {
            foreach (ExpectedScenario expected in ExpectedScenarios)
            {
                ReviewSnapshot snapshot =
                    SyntheticScenarioFactory.Create(expected.ScenarioId);

                TestAssert.Equal(expected.Seed, snapshot.Seed);
                TestAssert.Equal(expected.SourceCount, snapshot.Sources.Count);
                TestAssert.Equal(expected.PqmCount, snapshot.PqmRows.Count);
                TestAssert.Equal(
                    expected.PlanCheckCount,
                    snapshot.PlanCheckRows.Count);
                TestAssert.Equal(expected.FieldCount, snapshot.FieldRows.Count);
                TestAssert.Equal(
                    expected.MappingCount,
                    snapshot.StructureMappings.Count);
                TestAssert.Equal(
                    expected.DvhSeriesCount,
                    snapshot.DvhSeries.Count);

                AssertStatusCount(
                    snapshot.PqmRows.Select(row => row.Status),
                    ReviewStatusCodes.Pass,
                    expected.PqmPass);
                AssertStatusCount(
                    snapshot.PqmRows.Select(row => row.Status),
                    ReviewStatusCodes.Variation,
                    expected.PqmVariation);
                AssertStatusCount(
                    snapshot.PqmRows.Select(row => row.Status),
                    ReviewStatusCodes.Fail,
                    expected.PqmFail);
                AssertStatusCount(
                    snapshot.PlanCheckRows.Select(row => row.Status),
                    ReviewStatusCodes.Pass,
                    expected.PlanCheckPass);
                AssertStatusCount(
                    snapshot.PlanCheckRows.Select(row => row.Status),
                    ReviewStatusCodes.Variation,
                    expected.PlanCheckVariation);
                AssertStatusCount(
                    snapshot.PlanCheckRows.Select(row => row.Status),
                    ReviewStatusCodes.Fail,
                    expected.PlanCheckFail);
                AssertStatusCount(
                    snapshot.FieldRows.Select(row => row.IdStatus),
                    ReviewStatusCodes.Pass,
                    expected.FieldIdPass);
                AssertStatusCount(
                    snapshot.FieldRows.Select(row => row.IdStatus),
                    ReviewStatusCodes.Fail,
                    expected.FieldIdFail);
                AssertStatusCount(
                    snapshot.FieldRows.Select(row => row.NameStatus),
                    ReviewStatusCodes.Pass,
                    expected.FieldNamePass);
                AssertStatusCount(
                    snapshot.FieldRows.Select(row => row.NameStatus),
                    ReviewStatusCodes.Fail,
                    expected.FieldNameFail);
                AssertStatusCount(
                    snapshot.StructureMappings.Select(row => row.Status),
                    ReviewStatusCodes.Pass,
                    expected.MappingPass);
                AssertStatusCount(
                    snapshot.StructureMappings.Select(row => row.Status),
                    ReviewStatusCodes.Variation,
                    expected.MappingVariation);
                AssertStatusCount(
                    snapshot.Sources.Select(row => row.Status),
                    ReviewStatusCodes.Available,
                    expected.SourceAvailable);
                AssertStatusCount(
                    snapshot.Sources.Select(row => row.Status),
                    ReviewStatusCodes.Fallback,
                    expected.SourceFallback);

                TestAssert.Equal(
                    string.Join("|", RequiredStructures),
                    string.Join(
                        "|",
                        snapshot.DvhSeries.Select(series => series.StructureId)));
            }
        }

        public static void ProducesValidMonotoneDvhCurves()
        {
            foreach (ReviewSnapshot snapshot in SyntheticScenarioFactory.CreateAll())
            {
                ReviewSnapshotValidationResult validation =
                    ReviewSnapshotValidator.ValidateForSimulator(snapshot);
                TestAssert.True(
                    validation.IsValid,
                    snapshot.ScenarioId + ": " +
                    string.Join(
                        "; ",
                        validation.Issues.Select(issue =>
                            issue.Code + "@" + issue.Path)));

                foreach (ReviewDvhSeries series in snapshot.DvhSeries)
                {
                    TestAssert.True(
                        series.Points.Count >= 61,
                        snapshot.ScenarioId + "/" + series.StructureId +
                        " has too few analytical DVH points.");
                    TestAssert.Equal(0.0, series.Points[0].DoseGy);
                    TestAssert.Equal(100.0, series.Points[0].VolumePercent);
                    for (int index = 1; index < series.Points.Count; index++)
                    {
                        ReviewDvhPoint previous = series.Points[index - 1];
                        ReviewDvhPoint current = series.Points[index];
                        TestAssert.True(
                            current.DoseGy > previous.DoseGy,
                            "Dose in Gy must increase strictly.");
                        TestAssert.True(
                            current.VolumePercent <= previous.VolumePercent,
                            "Relative cumulative volume in percent must not increase.");
                        TestAssert.True(
                            current.VolumePercent >= 0.0 &&
                            current.VolumePercent <= 100.0);
                    }
                }
            }
        }

        public static void SerializesByteIdentically()
        {
            foreach (ExpectedScenario expected in ExpectedScenarios)
            {
                byte[] first = ReviewSnapshotJson.SerializeUtf8(
                    SyntheticScenarioFactory.Create(expected.ScenarioId));
                byte[] second = ReviewSnapshotJson.SerializeUtf8(
                    SyntheticScenarioFactory.Create(expected.ScenarioId));

                TestAssert.True(
                    first.SequenceEqual(second),
                    expected.ScenarioId + " did not serialize deterministically.");
            }
        }

        public static void ExposesVisibleDoseAndCheckFaults()
        {
            ReviewSnapshot baseline =
                SyntheticScenarioFactory.Create("baseline-pass");
            ReviewSnapshot target =
                SyntheticScenarioFactory.Create("target-underdose");
            ReviewSnapshot oar =
                SyntheticScenarioFactory.Create("oar-overdose");
            ReviewSnapshot metadata =
                SyntheticScenarioFactory.Create("metadata-plancheck");

            double baselineTargetAt57 = VolumeAtDose(
                baseline, "PTV_60", 57.0);
            double underTargetAt57 = VolumeAtDose(
                target, "PTV_60", 57.0);
            TestAssert.True(
                baselineTargetAt57 - underTargetAt57 >= 7.0,
                "The target-underdose DVH shift must be clearly visible.");
            TestAssert.Equal(
                1,
                target.PqmRows.Count(row =>
                    row.Status == ReviewStatusCodes.Variation));
            TestAssert.Equal(
                1,
                target.PqmRows.Count(row =>
                    row.Status == ReviewStatusCodes.Fail));

            double baselineCordAt40 = VolumeAtDose(
                baseline, "SpinalCord", 40.0);
            double overdoseCordAt40 = VolumeAtDose(
                oar, "SpinalCord", 40.0);
            TestAssert.True(
                overdoseCordAt40 - baselineCordAt40 >= 10.0,
                "The OAR-overdose DVH shift must be clearly visible.");
            TestAssert.Equal(
                2,
                oar.PqmRows.Count(row =>
                    row.Status == ReviewStatusCodes.Variation ||
                    row.Status == ReviewStatusCodes.Fail));

            TestAssert.True(
                metadata.PlanCheckRows.Count(row =>
                    !string.IsNullOrWhiteSpace(row.ObservedValue) &&
                    !string.IsNullOrWhiteSpace(row.ExpectedValue)) >= 5);
            TestAssert.Equal(
                4,
                metadata.PlanCheckRows.Count(row =>
                    row.Status == ReviewStatusCodes.Fail));
        }

        public static void PqmValuesAndStatusesMatchDvhCurves()
        {
            foreach (ReviewSnapshot snapshot in SyntheticScenarioFactory.CreateAll())
            {
                ReviewPqmRow nearMaximum = snapshot.PqmRows.Single(row =>
                    row.StableId == "pqm-cord-dmax");
                TestAssert.Equal("D0.1%", nearMaximum.Objective);
                TestAssert.True(
                    nearMaximum.Explanation.IndexOf(
                        "near-maximum",
                        StringComparison.OrdinalIgnoreCase) >= 0,
                    "The sampled hottest-volume metric must not be labeled as a point maximum.");

                foreach (ReviewPqmRow row in snapshot.PqmRows)
                {
                    ReviewDvhSeries series = snapshot.DvhSeries.Single(item =>
                        item.StructureId == row.ResolvedStructureId);
                    double calculated = CalculatePqmIndependently(
                        snapshot,
                        row,
                        series.Points);
                    double expectedDisplayValue = Math.Round(
                        calculated,
                        1,
                        MidpointRounding.AwayFromZero);

                    TestAssert.True(
                        row.AchievedValue.HasValue,
                        snapshot.ScenarioId + "/" + row.StableId +
                        " has no achieved value.");
                    TestAssert.True(
                        Math.Abs(
                            expectedDisplayValue -
                            row.AchievedValue.Value) <= 0.051,
                        string.Format(
                            "{0}/{1}: expected DVH-derived {2:0.0}, found {3:0.0}.",
                            snapshot.ScenarioId,
                            row.StableId,
                            expectedDisplayValue,
                            row.AchievedValue.Value));

                    string expectedStatus = EvaluateStatusIndependently(
                        expectedDisplayValue,
                        row.Comparator,
                        row.Goal,
                        row.Variation);
                    TestAssert.Equal(
                        expectedStatus,
                        row.Status,
                        string.Format(
                            "{0}/{1}: status must follow {2} goal {3} and variation {4}.",
                            snapshot.ScenarioId,
                            row.StableId,
                            row.Comparator,
                            row.Goal,
                            row.Variation));
                }
            }
        }

        public static void UsesFieldNamerAndShowsMappingDeviation()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("field-and-mapping");
            string[] expectedNames =
            {
                "179-181 T30 UZ",
                "179-181 UZa",
                "179-181 UZb",
                "179-181 GUZa",
                "179-181 GUZb"
            };

            TestAssert.Equal(
                string.Join("|", expectedNames),
                string.Join(
                    "|",
                    snapshot.FieldRows.Select(row => row.SuggestedName)));
            TestAssert.True(
                snapshot.FieldRows.All(row =>
                    row.IdStatus == ReviewStatusCodes.Pass &&
                    row.CurrentId == row.ExpectedId),
                "Every synthetic ID-okay row must remain a pass.");
            TestAssert.True(
                snapshot.FieldRows.All(row =>
                    row.NameStatus == ReviewStatusCodes.Fail &&
                    row.CurrentName != row.SuggestedName));
            TestAssert.Equal(
                1,
                snapshot.StructureMappings.Count(mapping =>
                    mapping.Status == ReviewStatusCodes.Variation &&
                    mapping.AvailableStructureIds.Count == 2));
        }

        public static void UsesNeutralOptionalFallbackOnly()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("optional-path-fallback");
            ReviewSourceStatus fallback = snapshot.Sources.Single(source =>
                source.Status == ReviewStatusCodes.Fallback);

            TestAssert.True(fallback.Optional);
            TestAssert.True(fallback.UsedFallback);
            TestAssert.Equal("Optional reference source", fallback.PathDisplayLabel);
            TestAssert.True(
                fallback.Message.IndexOf(
                    "embedded synthetic defaults",
                    StringComparison.OrdinalIgnoreCase) >= 0);
            AssertPublishSafeText(ReviewSnapshotJson.Serialize(snapshot));
        }

        public static void IntegratedMixedReviewCombinesVisibleFindings()
        {
            ReviewSnapshot snapshot =
                SyntheticScenarioFactory.Create("mixed-review");

            TestAssert.Equal(
                new DateTimeOffset(2026, 7, 30, 8, 6, 0, TimeSpan.Zero),
                snapshot.GeneratedUtc);
            TestAssert.Equal("Integrated mixed review", snapshot.ScenarioTitle);
            TestAssert.Equal(
                "56.3|90.5|62.0|60.9|32.9|49.6",
                string.Join(
                    "|",
                    snapshot.PqmRows.Select(row =>
                        row.AchievedValue.Value.ToString("0.0",
                            System.Globalization.CultureInfo.InvariantCulture))));
            TestAssert.True(
                snapshot.PlanCheckRows
                    .Where(row => row.Status != ReviewStatusCodes.Pass)
                    .All(row => row.Message.StartsWith(
                        "Injected synthetic finding:",
                        StringComparison.Ordinal)));
            TestAssert.True(snapshot.FieldRows.All(row =>
                row.IdStatus == ReviewStatusCodes.Pass &&
                row.NameStatus == ReviewStatusCodes.Fail));
            TestAssert.Equal(
                "179-181 T30 UZ|179-181 UZa|179-181 UZb|179-181 GUZa|179-181 GUZb",
                string.Join(
                    "|",
                    snapshot.FieldRows.Select(row => row.SuggestedName)));
            TestAssert.Equal(
                1,
                snapshot.StructureMappings.Count(mapping =>
                    mapping.Status == ReviewStatusCodes.Variation));
            TestAssert.Equal(
                1,
                snapshot.Sources.Count(source =>
                    source.Status == ReviewStatusCodes.Fallback &&
                    source.Optional));
        }

        public static void ContainsSyntheticProvenanceOnly()
        {
            foreach (ReviewSnapshot snapshot in SyntheticScenarioFactory.CreateAll())
            {
                TestAssert.True(snapshot.Synthetic);
                TestAssert.Equal(SyntheticWatermark, snapshot.Report.Watermark);
                TestAssert.Equal(SyntheticWatermark, snapshot.Report.ModeLabel);
                TestAssert.True(
                    snapshot.ProvenanceText.IndexOf(
                        "analytically generated",
                        StringComparison.OrdinalIgnoreCase) >= 0);
                TestAssert.Equal(
                    "Synthetic demonstration",
                    snapshot.PatientDisplayLabel);
                AssertPublishSafeText(ReviewSnapshotJson.Serialize(snapshot));
            }
        }

        public static void CheckedInJsonMatchesFactory()
        {
            string scenarioDirectory = FindScenarioDirectory();
            foreach (ExpectedScenario expected in ExpectedScenarios)
            {
                string path = Path.Combine(
                    scenarioDirectory,
                    expected.ScenarioId + ".json");
                TestAssert.True(File.Exists(path), "Missing " + path);
                TestAssert.True(
                    File.ReadAllBytes(path).SequenceEqual(
                        ReviewSnapshotJson.SerializeUtf8(
                            SyntheticScenarioFactory.Create(expected.ScenarioId))),
                    expected.ScenarioId + ".json differs from the factory output.");
            }

            foreach (string path in Directory.GetFiles(scenarioDirectory))
            {
                string extension = Path.GetExtension(path);
                if (extension.Equals(".json", StringComparison.OrdinalIgnoreCase) ||
                    extension.Equals(".md", StringComparison.OrdinalIgnoreCase))
                {
                    AssertPublishSafeText(File.ReadAllText(path, Encoding.UTF8));
                }
            }
        }

        public static int ExportScenarios(string outputDirectory)
        {
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                outputDirectory = FindScenarioDirectory();
            }

            Directory.CreateDirectory(outputDirectory);
            foreach (ReviewSnapshot snapshot in SyntheticScenarioFactory.CreateAll())
            {
                File.WriteAllBytes(
                    Path.Combine(outputDirectory, snapshot.ScenarioId + ".json"),
                    ReviewSnapshotJson.SerializeUtf8(snapshot));
            }

            Console.WriteLine(
                "Exported {0} deterministic scenarios to {1}.",
                ExpectedScenarios.Length,
                Path.GetFullPath(outputDirectory));
            return 0;
        }

        private static void AssertStatusCount(
            IEnumerable<string> statuses,
            string status,
            int expected)
        {
            TestAssert.Equal(expected, statuses.Count(value => value == status));
        }

        private static double VolumeAtDose(
            ReviewSnapshot snapshot,
            string structureId,
            double doseGy)
        {
            return snapshot.DvhSeries
                .Single(series => series.StructureId == structureId)
                .Points
                .Single(point => Math.Abs(point.DoseGy - doseGy) < 0.0001)
                .VolumePercent;
        }

        private static double CalculatePqmIndependently(
            ReviewSnapshot snapshot,
            ReviewPqmRow row,
            IList<ReviewDvhPoint> points)
        {
            switch (row.Objective)
            {
                case "D95":
                    return DoseAtVolumeIndependently(points, 95.0);
                case "V95":
                    ReviewPlanRow plan = snapshot.Plans.Single(item =>
                        item.PlanKey == snapshot.ActivePlanKey);
                    TestAssert.True(plan.TotalDoseGy.HasValue);
                    return VolumeAtDoseIndependently(
                        points,
                        plan.TotalDoseGy.Value * 0.95);
                case "D2":
                    return DoseAtVolumeIndependently(points, 2.0);
                case "D0.1%":
                    return DoseAtVolumeIndependently(points, 0.1);
                case "Dmean":
                    return MeanDoseIndependently(points);
                case "V30":
                    return VolumeAtDoseIndependently(points, 30.0);
                default:
                    throw new InvalidOperationException(
                        "No independent synthetic DVH metric for " +
                        row.Objective + ".");
            }
        }

        private static double VolumeAtDoseIndependently(
            IList<ReviewDvhPoint> points,
            double doseGy)
        {
            if (doseGy < points[0].DoseGy ||
                doseGy > points[points.Count - 1].DoseGy)
            {
                throw new InvalidOperationException(
                    "Requested Vx dose is outside the sampled DVH.");
            }

            for (int index = 1; index < points.Count; index++)
            {
                ReviewDvhPoint lower = points[index - 1];
                ReviewDvhPoint upper = points[index];
                if (doseGy <= upper.DoseGy)
                {
                    double fraction =
                        (doseGy - lower.DoseGy) /
                        (upper.DoseGy - lower.DoseGy);
                    return lower.VolumePercent +
                           fraction *
                           (upper.VolumePercent - lower.VolumePercent);
                }
            }

            return points[points.Count - 1].VolumePercent;
        }

        private static double DoseAtVolumeIndependently(
            IList<ReviewDvhPoint> points,
            double volumePercent)
        {
            if (volumePercent > points[0].VolumePercent ||
                volumePercent < points[points.Count - 1].VolumePercent)
            {
                throw new InvalidOperationException(
                    "Requested Dx volume is outside the sampled DVH.");
            }

            for (int index = 1; index < points.Count; index++)
            {
                ReviewDvhPoint lowerDose = points[index - 1];
                ReviewDvhPoint upperDose = points[index];
                if (volumePercent >= upperDose.VolumePercent)
                {
                    double volumeDifference =
                        lowerDose.VolumePercent -
                        upperDose.VolumePercent;
                    if (volumeDifference <= 0.0)
                    {
                        return upperDose.DoseGy;
                    }

                    double fraction =
                        (lowerDose.VolumePercent - volumePercent) /
                        volumeDifference;
                    return lowerDose.DoseGy +
                           fraction *
                           (upperDose.DoseGy - lowerDose.DoseGy);
                }
            }

            return points[points.Count - 1].DoseGy;
        }

        private static double MeanDoseIndependently(
            IList<ReviewDvhPoint> points)
        {
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

        private static string EvaluateStatusIndependently(
            double achieved,
            string comparator,
            double? goal,
            double? variation)
        {
            TestAssert.True(goal.HasValue);
            TestAssert.True(variation.HasValue);
            if (comparator == ">=")
            {
                if (achieved >= goal.Value)
                {
                    return ReviewStatusCodes.Pass;
                }

                return achieved >= variation.Value
                    ? ReviewStatusCodes.Variation
                    : ReviewStatusCodes.Fail;
            }

            if (comparator == "<=")
            {
                if (achieved <= goal.Value)
                {
                    return ReviewStatusCodes.Pass;
                }

                return achieved <= variation.Value
                    ? ReviewStatusCodes.Variation
                    : ReviewStatusCodes.Fail;
            }

            throw new InvalidOperationException(
                "Unsupported synthetic PQM comparator: " + comparator);
        }

        private static void AssertPublishSafeText(string text)
        {
            string[] privateTokens =
            {
                "medizin" + "." + "uni-" + "leipzig",
                "u" + "ke",
                "workgroup",
                "archiv",
                "private-user-name",
                "private-clinical-host",
                "patientname",
                "patientid",
                "accession",
                "medical record",
                "\"mrn\"",
                "birthdate",
                "birth date",
                "date of birth",
                "geburtsdatum"
            };
            foreach (string token in privateTokens)
            {
                TestAssert.True(
                    text.IndexOf(token, StringComparison.OrdinalIgnoreCase) < 0,
                    "Publish-safe scenario text contains forbidden token: " + token);
            }

            TestAssert.False(
                Regex.IsMatch(
                    text,
                    @"(?<![\d.])\d+(?:\.\d+){4,}(?![\d.])",
                    RegexOptions.CultureInvariant),
                "Publish-safe scenario text contains a DICOM UID pattern.");
            TestAssert.False(
                Regex.IsMatch(
                    text,
                    @"(?<![A-Za-z0-9])(?:[A-Za-z]:[\\/]|\\\\)",
                    RegexOptions.CultureInvariant),
                "Publish-safe scenario text contains an absolute path.");
        }

        private static string FindScenarioDirectory()
        {
            DirectoryInfo current = new DirectoryInfo(
                Environment.CurrentDirectory);
            while (current != null)
            {
                string solution = Path.Combine(current.FullName, "ClearPlan.sln");
                if (File.Exists(solution))
                {
                    return Path.Combine(
                        current.FullName,
                        "ClearPlan.Simulator",
                        "Scenarios");
                }

                current = current.Parent;
            }

            throw new DirectoryNotFoundException(
                "Could not locate ClearPlan.Simulator/Scenarios.");
        }

        private sealed class ExpectedScenario
        {
            public ExpectedScenario(
                string scenarioId,
                int seed,
                int sourceCount,
                int pqmCount,
                int planCheckCount,
                int fieldCount,
                int mappingCount,
                int dvhSeriesCount,
                int pqmPass,
                int pqmVariation,
                int pqmFail,
                int planCheckPass,
                int planCheckVariation,
                int planCheckFail,
                int fieldIdPass,
                int fieldIdFail,
                int fieldNamePass,
                int fieldNameFail,
                int mappingPass,
                int mappingVariation,
                int sourceAvailable,
                int sourceFallback)
            {
                ScenarioId = scenarioId;
                Seed = seed;
                SourceCount = sourceCount;
                PqmCount = pqmCount;
                PlanCheckCount = planCheckCount;
                FieldCount = fieldCount;
                MappingCount = mappingCount;
                DvhSeriesCount = dvhSeriesCount;
                PqmPass = pqmPass;
                PqmVariation = pqmVariation;
                PqmFail = pqmFail;
                PlanCheckPass = planCheckPass;
                PlanCheckVariation = planCheckVariation;
                PlanCheckFail = planCheckFail;
                FieldIdPass = fieldIdPass;
                FieldIdFail = fieldIdFail;
                FieldNamePass = fieldNamePass;
                FieldNameFail = fieldNameFail;
                MappingPass = mappingPass;
                MappingVariation = mappingVariation;
                SourceAvailable = sourceAvailable;
                SourceFallback = sourceFallback;
            }

            public string ScenarioId { get; private set; }
            public int Seed { get; private set; }
            public int SourceCount { get; private set; }
            public int PqmCount { get; private set; }
            public int PlanCheckCount { get; private set; }
            public int FieldCount { get; private set; }
            public int MappingCount { get; private set; }
            public int DvhSeriesCount { get; private set; }
            public int PqmPass { get; private set; }
            public int PqmVariation { get; private set; }
            public int PqmFail { get; private set; }
            public int PlanCheckPass { get; private set; }
            public int PlanCheckVariation { get; private set; }
            public int PlanCheckFail { get; private set; }
            public int FieldIdPass { get; private set; }
            public int FieldIdFail { get; private set; }
            public int FieldNamePass { get; private set; }
            public int FieldNameFail { get; private set; }
            public int MappingPass { get; private set; }
            public int MappingVariation { get; private set; }
            public int SourceAvailable { get; private set; }
            public int SourceFallback { get; private set; }
        }
    }
}
