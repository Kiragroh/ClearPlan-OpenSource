using System;
using System.Linq;
using System.Text;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewSnapshotTests
    {
        public static void AcceptsValidMinimalSnapshot()
        {
            const string json =
                "{\"schemaVersion\":1,\"scenarioId\":\"minimal-pass\",\"synthetic\":true," +
                "\"activePlanKey\":\"plan-1\",\"plans\":[{\"planKey\":\"plan-1\"," +
                "\"displayLabel\":\"Synthetic plan\"}]}";

            ReviewSnapshot snapshot = ReviewSnapshotJson.Deserialize(json);
            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(result.IsValid, JoinIssues(result));
            TestAssert.NotNull(snapshot.Sources);
            TestAssert.NotNull(snapshot.PqmRows);
            TestAssert.NotNull(snapshot.PlanCheckRows);
            TestAssert.NotNull(snapshot.FieldRows);
            TestAssert.NotNull(snapshot.StructureMappings);
            TestAssert.NotNull(snapshot.DvhSeries);
            TestAssert.NotNull(snapshot.Report);
        }

        public static void RejectsUnsupportedSchemaVersion()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.SchemaVersion = ReviewSnapshot.CurrentSchemaVersion + 1;

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.UnsupportedSchemaVersion),
                JoinIssues(result));
        }

        public static void RejectsNonSyntheticSimulatorInput()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.Synthetic = false;

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.SimulatorRequiresSynthetic),
                JoinIssues(result));
        }

        public static void RejectsDuplicateStableIds()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.PqmRows.Add(new ReviewPqmRow
            {
                StableId = snapshot.PqmRows[0].StableId,
                TemplateStructure = "Target",
                ResolvedStructureId = "PTV_1",
                Objective = "D95%",
                Status = ReviewStatusCodes.Pass,
                Severity = ReviewSeverityCodes.None
            });

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.DuplicateStableId),
                JoinIssues(result));
        }

        public static void RejectsMissingActivePlan()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.ActivePlanKey = "missing-plan";

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.MissingActivePlan),
                JoinIssues(result));
        }

        public static void RejectsInvalidStatusAndSeverity()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.PqmRows[0].Status = "mystery";
            snapshot.PqmRows[0].Severity = "critical";

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.InvalidStatus),
                JoinIssues(result));
            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.InvalidSeverity),
                JoinIssues(result));
        }

        public static void RejectsNonFiniteDose()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.DvhSeries[0].Points[1].DoseGy = double.NaN;

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.InvalidDose),
                JoinIssues(result));
        }

        public static void RequiresIncreasingDoseOrder()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.DvhSeries[0].Points.Add(new ReviewDvhPoint(10.0, 90.0));

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.DecreasingDose),
                JoinIssues(result));
        }

        public static void RejectsIncreasingCumulativeVolume()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.DvhSeries[0].Points[0].VolumePercent = 90.0;
            snapshot.DvhSeries[0].Points[1].VolumePercent = 95.0;

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.IncreasingCumulativeVolume),
                JoinIssues(result));
        }

        public static void RejectsMissingReferencedStructure()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.StructureMappings[0].SelectedStructureId = "MISSING";

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                HasIssue(result, ReviewSnapshotValidationCodes.MissingReferencedStructure),
                JoinIssues(result));
        }

        public static void RejectsIdentityBearingClinicalContent()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.PatientDisplayLabel = @"\\private-server\patient-share";
            snapshot.PlanDisplayLabel = "1.2.840.10008.5.1.4.1.1.481.5";

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(
                result.Issues.Count(issue =>
                    issue.Code == ReviewSnapshotValidationCodes.UnsafeDisplayLabel) >= 2,
                JoinIssues(result));
        }

        public static void RejectsEmbeddedPrivateContentAcrossAllTextSurfaces()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            string pathText = @"Source: C:\synthetic-private\case.json";
            string uncText = @"Source: \\synthetic-private\review\case.json";
            string uidText = "Plan 1.2.840.10008.5.1.4.1.1.481.5";

            snapshot.ScenarioId = pathText;
            snapshot.ScenarioTitle = uidText;
            snapshot.ScenarioDescription = pathText;
            snapshot.ProvenanceText = uncText;
            snapshot.PatientDisplayLabel = pathText;
            snapshot.PlanDisplayLabel = uidText;
            snapshot.ActivePlanKey = pathText;

            ReviewSourceStatus source = snapshot.Sources[0];
            source.StableId = uidText;
            source.SourceCode = pathText;
            source.SourceType = uidText;
            source.PathDisplayLabel = pathText;
            source.Message = uncText;

            ReviewPlanRow plan = snapshot.Plans[0];
            plan.PlanKey = pathText;
            plan.DisplayLabel = uidText;
            plan.TargetDisplayLabel = uncText;

            ReviewPqmRow pqm = snapshot.PqmRows[0];
            pqm.StableId = pathText;
            pqm.TemplateCode = uidText;
            pqm.TemplateStructure = pathText;
            pqm.ResolvedStructureId = uidText;
            pqm.StructureOptions[0] = uncText;
            pqm.Objective = pathText;
            pqm.Comparator = uidText;
            pqm.Explanation = uncText;

            ReviewCheckRow check = snapshot.PlanCheckRows[0];
            check.CheckCode = pathText;
            check.Category = uidText;
            check.ObservedValue = pathText;
            check.ExpectedValue = uidText;
            check.Message = uncText;

            ReviewFieldRow field = snapshot.FieldRows[0];
            field.StableId = pathText;
            field.CurrentId = uidText;
            field.ExpectedId = pathText;
            field.CurrentName = uncText;
            field.SuggestedName = uidText;

            ReviewStructureMapping mapping = snapshot.StructureMappings[0];
            mapping.StableId = uidText;
            mapping.TemplateStructure = pathText;
            mapping.SelectedStructureId = uidText;
            mapping.AvailableStructureIds[0] = uncText;
            mapping.Message = pathText;

            ReviewDvhSeries series = snapshot.DvhSeries[0];
            series.StableId = pathText;
            series.StructureId = uidText;
            series.DisplayName = uncText;
            series.ColorHex = uidText;

            snapshot.Report.Title = pathText;
            snapshot.Report.Subtitle = uidText;
            snapshot.Report.ModeLabel = uncText;
            snapshot.Report.Watermark = pathText;
            snapshot.Report.OutputFileLabel = uidText;
            snapshot.Report.Notes.Add(uncText);

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            string[] expectedUnsafePaths =
            {
                "$.scenarioId",
                "$.scenarioTitle",
                "$.scenarioDescription",
                "$.provenanceText",
                "$.patientDisplayLabel",
                "$.planDisplayLabel",
                "$.activePlanKey",
                "$.sources[0].stableId",
                "$.sources[0].sourceCode",
                "$.sources[0].sourceType",
                "$.sources[0].pathDisplayLabel",
                "$.sources[0].message",
                "$.plans[0].planKey",
                "$.plans[0].displayLabel",
                "$.plans[0].targetDisplayLabel",
                "$.pqmRows[0].stableId",
                "$.pqmRows[0].templateCode",
                "$.pqmRows[0].templateStructure",
                "$.pqmRows[0].resolvedStructureId",
                "$.pqmRows[0].structureOptions[0]",
                "$.pqmRows[0].objective",
                "$.pqmRows[0].comparator",
                "$.pqmRows[0].explanation",
                "$.planCheckRows[0].checkCode",
                "$.planCheckRows[0].category",
                "$.planCheckRows[0].observedValue",
                "$.planCheckRows[0].expectedValue",
                "$.planCheckRows[0].message",
                "$.fieldRows[0].stableId",
                "$.fieldRows[0].currentId",
                "$.fieldRows[0].expectedId",
                "$.fieldRows[0].currentName",
                "$.fieldRows[0].suggestedName",
                "$.structureMappings[0].stableId",
                "$.structureMappings[0].templateStructure",
                "$.structureMappings[0].selectedStructureId",
                "$.structureMappings[0].availableStructureIds[0]",
                "$.structureMappings[0].message",
                "$.dvhSeries[0].stableId",
                "$.dvhSeries[0].structureId",
                "$.dvhSeries[0].displayName",
                "$.dvhSeries[0].colorHex",
                "$.report.title",
                "$.report.subtitle",
                "$.report.modeLabel",
                "$.report.watermark",
                "$.report.outputFileLabel",
                "$.report.notes[0]"
            };

            foreach (string expectedPath in expectedUnsafePaths)
            {
                TestAssert.True(
                    result.Issues.Any(issue =>
                        issue.Code ==
                            ReviewSnapshotValidationCodes.UnsafeDisplayLabel &&
                        issue.Path == expectedPath),
                    "Expected an unsafe-content issue at " + expectedPath +
                    Environment.NewLine + JoinIssues(result));
            }
        }

        public static void AllowsGenericSyntheticTextAcrossTheContract()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();
            snapshot.ScenarioDescription =
                "Demonstration of target coverage and an optional source warning.";
            snapshot.ProvenanceText =
                "Analytical values generated from a fixed seed.";
            snapshot.Sources[0].Message =
                "Optional source unavailable; embedded fallback active.";
            snapshot.PlanCheckRows[0].Message =
                "Synthetic prescription and metadata are complete.";
            snapshot.Report.Notes.Add(
                "Values are intended for software demonstration only.");

            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);

            TestAssert.True(result.IsValid, JoinIssues(result));
        }

        public static void JsonRoundTripIsDeterministic()
        {
            ReviewSnapshot snapshot = CreateValidSnapshot();

            string first = ReviewSnapshotJson.Serialize(snapshot);
            byte[] firstUtf8 = ReviewSnapshotJson.SerializeUtf8(snapshot);
            byte[] secondUtf8 = ReviewSnapshotJson.SerializeUtf8(snapshot);
            string roundTrip = ReviewSnapshotJson.Serialize(
                ReviewSnapshotJson.Deserialize(first));

            TestAssert.True(firstUtf8.SequenceEqual(secondUtf8));
            TestAssert.Equal(first, roundTrip);
            TestAssert.False(
                firstUtf8.Length >= 3 &&
                firstUtf8[0] == 0xEF &&
                firstUtf8[1] == 0xBB &&
                firstUtf8[2] == 0xBF,
                "UTF-8 output must not contain a byte-order mark.");
            TestAssert.Equal(first, new UTF8Encoding(false, true).GetString(firstUtf8));
            TestAssert.True(
                first.IndexOf("\"schemaVersion\"", StringComparison.Ordinal) <
                first.IndexOf("\"scenarioId\"", StringComparison.Ordinal));
            TestAssert.True(
                first.IndexOf("\"plans\"", StringComparison.Ordinal) <
                first.IndexOf("\"pqmRows\"", StringComparison.Ordinal));
            TestAssert.True(
                first.IndexOf("\"pqmRows\"", StringComparison.Ordinal) <
                first.IndexOf("\"dvhSeries\"", StringComparison.Ordinal));
        }

        private static ReviewSnapshot CreateValidSnapshot()
        {
            var snapshot = new ReviewSnapshot
            {
                SchemaVersion = ReviewSnapshot.CurrentSchemaVersion,
                ScenarioId = "baseline-pass",
                ScenarioTitle = "Baseline pass",
                ScenarioDescription = "Deterministic synthetic review case.",
                Seed = 4101,
                Synthetic = true,
                GeneratedUtc = new DateTimeOffset(2026, 7, 30, 8, 0, 0, TimeSpan.Zero),
                PatientDisplayLabel = "Synthetic patient",
                PlanDisplayLabel = "Synthetic plan A",
                ProvenanceText = "Analytically generated test values.",
                ActivePlanKey = "plan-1"
            };

            snapshot.Sources.Add(new ReviewSourceStatus
            {
                StableId = "source-synthetic",
                SourceCode = "synthetic",
                SourceType = "analytical",
                Status = ReviewStatusCodes.Available,
                PathDisplayLabel = "Embedded scenario",
                Message = "Available"
            });
            snapshot.Plans.Add(new ReviewPlanRow
            {
                PlanKey = "plan-1",
                DisplayLabel = "Synthetic plan A",
                DosePerFractionGy = 2.0,
                TotalDoseGy = 20.0,
                FractionCount = 10,
                TargetDisplayLabel = "PTV_1",
                Status = ReviewStatusCodes.Pass
            });
            snapshot.PqmRows.Add(new ReviewPqmRow
            {
                StableId = "pqm-target-d95",
                TemplateCode = "target",
                TemplateStructure = "Target",
                ResolvedStructureId = "PTV_1",
                Objective = "D95%",
                Goal = 19.0,
                Variation = 18.0,
                AchievedValue = 19.4,
                Unit = ReviewUnitCodes.Gray,
                Status = ReviewStatusCodes.Pass,
                Severity = ReviewSeverityCodes.None,
                Explanation = "Synthetic target coverage is conforming."
            });
            snapshot.PqmRows[0].StructureOptions.Add("PTV_1");
            snapshot.PlanCheckRows.Add(new ReviewCheckRow
            {
                CheckCode = "plan.prescription",
                Category = "Prescription",
                Status = ReviewStatusCodes.Pass,
                Severity = ReviewSeverityCodes.None,
                ObservedValue = "10 x 2 Gy",
                ExpectedValue = "10 x 2 Gy",
                Unit = ReviewUnitCodes.Text,
                Message = "Prescription is complete."
            });
            snapshot.FieldRows.Add(new ReviewFieldRow
            {
                StableId = "field-1",
                TreatmentOrder = 1,
                BeamNumber = 1,
                CurrentId = "1",
                ExpectedId = "1",
                CurrentName = "181-179 UZ",
                SuggestedName = "181-179 UZ",
                IdStatus = ReviewStatusCodes.Pass,
                NameStatus = ReviewStatusCodes.Pass
            });
            snapshot.StructureMappings.Add(new ReviewStructureMapping
            {
                StableId = "mapping-target",
                TemplateStructure = "Target",
                SelectedStructureId = "PTV_1"
            });
            snapshot.StructureMappings[0].AvailableStructureIds.Add("PTV_1");
            var targetDvh = new ReviewDvhSeries
            {
                StableId = "dvh-ptv-1",
                StructureId = "PTV_1",
                DisplayName = "PTV_1",
                Role = ReviewDvhRoleCodes.Target,
                ColorHex = "#315DDC",
                LineStyle = ReviewLineStyleCodes.Solid,
                Selected = true,
                VolumeCc = 125.0
            };
            targetDvh.Points.Add(new ReviewDvhPoint(0.0, 100.0));
            targetDvh.Points.Add(new ReviewDvhPoint(20.0, 95.0));
            snapshot.DvhSeries.Add(targetDvh);
            snapshot.Report.Title = "ClearPlan synthetic review";
            snapshot.Report.ModeLabel = "Synthetic demonstration";
            snapshot.Report.Watermark = "SYNTHETIC DEMONSTRATION - NOT FOR CLINICAL USE";
            snapshot.Report.OutputFileLabel = "baseline-pass.pdf";

            return snapshot;
        }

        private static bool HasIssue(
            ReviewSnapshotValidationResult result,
            string code)
        {
            return result.Issues.Any(issue => issue.Code == code);
        }

        private static string JoinIssues(ReviewSnapshotValidationResult result)
        {
            return string.Join(
                Environment.NewLine,
                result.Issues.Select(issue =>
                    string.Format("{0} {1}: {2}", issue.Code, issue.Path, issue.Message)));
        }
    }
}
