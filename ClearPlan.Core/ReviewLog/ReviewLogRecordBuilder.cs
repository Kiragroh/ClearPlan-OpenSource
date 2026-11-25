using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.ReviewLog
{
    /// <summary>All rows of one review, keyed by table name, already formatted as CSV field strings.</summary>
    public sealed class ReviewLogRecord
    {
        public ReviewLogRecord()
        {
            Rows = new Dictionary<string, List<string[]>>(StringComparer.Ordinal);
            foreach (string table in ReviewLogSchema.WriteOrder) Rows[table] = new List<string[]>();
        }

        public Guid ReviewId { get; set; }
        public Dictionary<string, List<string[]>> Rows { get; private set; }
    }

    /// <summary>
    /// Projects a detached <see cref="ReviewSnapshot"/> plus its <see cref="ReviewLogContext"/> into the five log tables.
    /// Pure and portable: no ESAPI, no file IO. Unavailable values stay empty; they are never written as 0.
    /// </summary>
    public static class ReviewLogRecordBuilder
    {
        public const string EclipseGoalSourcePrefix = "Eclipse Clinical Goals";

        public static ReviewLogRecord Build(ReviewSnapshot snapshot, ReviewLogContext context,
            ReviewLogPseudonymizer pseudonymizer, ReviewLogOptions options)
        {
            if (context == null) throw new ArgumentNullException("context");
            if (pseudonymizer == null) throw new ArgumentNullException("pseudonymizer");
            if (!ReviewLogModes.IsKnown(context.Mode)) throw new ArgumentException("Unknown review-log mode '" + context.Mode + "'.", "context");
            options = options ?? new ReviewLogOptions();
            var scrubber = new ReviewLogTextScrubber(context.SensitiveTerms);
            Func<string, string> text = value => options.LogMessages ? scrubber.Scrub(value) : string.Empty;
            string id = context.ReviewId.ToString("D");
            var record = new ReviewLogRecord { ReviewId = context.ReviewId };

            // ruleset_revision
            foreach (ReviewLogConfigurationInput input in (context.ConfigurationInputs ?? new List<ReviewLogConfigurationInput>())
                .Where(i => i != null).OrderBy(i => i.ConfigKey, StringComparer.Ordinal))
            {
                record.Rows[ReviewLogSchema.RulesetRevisionTable].Add(new[]
                {
                    id, input.ConfigKey ?? string.Empty, input.Category ?? string.Empty, ReviewLogCsv.Integer(input.RevisionNumber),
                    (input.Sha256 ?? string.Empty).ToLowerInvariant(), ReviewLogCsv.Utc(input.RevisionCreatedUtc), input.Availability ?? string.Empty
                });
            }

            // goal_result
            var goalStatuses = new List<string>();
            var tableDefinitionLines = new List<string>();
            if (snapshot != null)
            {
                int rowNo = 0;
                foreach (ReviewPqmRow row in snapshot.PqmRows.Where(r => r != null))
                {
                    rowNo++;
                    string status = ReviewLogGoalStatus.FromSnapshot(row.Status);
                    goalStatuses.Add(status);
                    string sourceKind = SourceKind(row.SourceLabel);
                    if (sourceKind == "constraint_table")
                        tableDefinitionLines.Add(string.Join("\u001f", new[] { row.TemplateCode, row.TemplateStructure, row.Objective,
                            row.Comparator, ReviewLogCsv.Number(row.Goal), ReviewLogCsv.Number(row.Variation), row.Unit }.Select(v => v ?? string.Empty)));
                    record.Rows[ReviewLogSchema.GoalResultTable].Add(new[]
                    {
                        id, rowNo.ToString(CultureInfo.InvariantCulture), row.StableId ?? string.Empty, sourceKind,
                        row.TemplateCode ?? string.Empty, row.TemplateStructure ?? string.Empty, row.ResolvedStructureId ?? string.Empty,
                        MappingStatus(snapshot, row), row.Objective ?? string.Empty, row.Comparator ?? string.Empty,
                        ReviewLogCsv.Number(row.Goal), ReviewLogCsv.Number(row.Variation), ReviewLogCsv.Number(row.AchievedValue),
                        row.Unit ?? string.Empty, status, string.Empty, row.SourceLabel ?? string.Empty
                    });
                }
            }

            // plan_parameter
            if (snapshot != null) AddParameters(record, id, snapshot.PlanAnalysis, text);

            // check_finding (included rows and, if known, excluded rows; OK rows are kept for denominators)
            int checks = 0, open = 0;
            if (snapshot != null)
            {
                int rowNo = 0;
                foreach (ReviewCheckRow row in snapshot.PlanCheckRows.Where(r => r != null))
                {
                    string status = ReviewLogCheckStatus.FromSnapshot(row.Status);
                    checks++;
                    if (ReviewLogCheckStatus.IsOpen(status)) open++;
                    record.Rows[ReviewLogSchema.CheckFindingTable].Add(CheckRow(id, ++rowNo, row, status, true, text));
                }
                foreach (ReviewCheckRow row in (context.ExcludedCheckRows ?? new List<ReviewCheckRow>()).Where(r => r != null))
                    record.Rows[ReviewLogSchema.CheckFindingTable].Add(CheckRow(id, ++rowNo, row, ReviewLogCheckStatus.FromSnapshot(row.Status), false, text));
            }

            // review (always last)
            string goalTableSha = tableDefinitionLines.Count == 0 ? string.Empty
                : ReviewLogRulesetResolver.Sha256Hex(string.Join("\n", tableDefinitionLines));
            string rulesetSha = RulesetSha256(context.ConfigurationInputs, goalTableSha);
            ReviewPlanRow plan = snapshot == null ? null
                : snapshot.Plans.FirstOrDefault(p => p != null && p.PlanKey == snapshot.ActivePlanKey) ?? snapshot.Plans.FirstOrDefault(p => p != null);
            List<ReviewBeamAnalysis> beams = snapshot == null || snapshot.PlanAnalysis == null ? new List<ReviewBeamAnalysis>()
                : snapshot.PlanAnalysis.Beams.Where(b => b != null).ToList();
            int? excluded = snapshot == null ? (int?)null
                : Math.Max(snapshot.DisabledCheckCount, (context.ExcludedCheckRows ?? new List<ReviewCheckRow>()).Count);
            Func<string, string> count = s => snapshot == null ? string.Empty
                : goalStatuses.Count(g => g == s).ToString(CultureInfo.InvariantCulture);
            record.Rows[ReviewLogSchema.ReviewTable].Add(new[]
            {
                ReviewLogSchema.SchemaVersion.ToString(CultureInfo.InvariantCulture), id, context.Mode,
                ReviewLogCsv.Utc(context.StartedUtc), ReviewLogCsv.Utc(context.FinishedUtc ?? DateTimeOffset.UtcNow),
                context.ClearPlanVersion ?? string.Empty, context.BuildId ?? string.Empty, pseudonymizer.IdentifierMode,
                pseudonymizer.User(context.UserId), pseudonymizer.Patient(context.PatientId),
                pseudonymizer.Course(context.PatientId, context.CourseId),
                pseudonymizer.Plan(context.PlanUid, context.PatientId, context.CourseId, context.PlanId),
                options.LogPlanLabel || pseudonymizer.ClearIdentifiers ? (context.PlanId ?? string.Empty) : string.Empty,
                context.PlanningItemType ?? string.Empty,
                ReviewLogCsv.Utc(context.PlanCreatedUtc ?? (plan == null ? null : plan.CreatedUtc)),
                context.PlanApprovalStatus ?? string.Empty,
                ReviewLogCsv.Utc(context.PlanningApprovalUtc), ReviewLogCsv.Utc(context.TreatmentApprovalUtc),
                Distinct(beams.Select(b => b.MachineId)), Distinct(beams.Select(b => b.Technique)), Distinct(beams.Select(b => b.MlcModel)),
                ReviewLogCsv.Number(plan == null ? null : plan.TotalDoseGy),
                ReviewLogCsv.Number(plan == null ? null : plan.DosePerFractionGy),
                ReviewLogCsv.Integer(plan == null ? null : plan.FractionCount),
                context.GoalTableId ?? string.Empty, goalTableSha, rulesetSha,
                snapshot == null ? string.Empty : goalStatuses.Count.ToString(CultureInfo.InvariantCulture),
                count(ReviewLogGoalStatus.Goal), count(ReviewLogGoalStatus.Variation), count(ReviewLogGoalStatus.NotMet),
                count(ReviewLogGoalStatus.NotEvaluated),
                snapshot == null ? string.Empty : checks.ToString(CultureInfo.InvariantCulture),
                ReviewLogCsv.Integer(excluded),
                snapshot == null ? string.Empty : open.ToString(CultureInfo.InvariantCulture),
                ReviewLogCsv.Flag(context.ReportWritten), (context.ReportSha256 ?? string.Empty).ToLowerInvariant(),
                ReviewLogCsv.Flag(context.ReportFiled), context.RunStatus ?? ReviewLogRunStatus.Ok,
                scrubber.Scrub(context.RunMessage)
            });
            return record;
        }

        public static string SourceKind(string sourceLabel)
        {
            if (string.IsNullOrWhiteSpace(sourceLabel)) return "constraint_table";
            if (sourceLabel.StartsWith(EclipseGoalSourcePrefix, StringComparison.OrdinalIgnoreCase)) return "eclipse_clinical_goal";
            if (string.Equals(sourceLabel, UserTargetCoverageRule.SourceLabel, StringComparison.Ordinal)) return "default_rule";
            return "constraint_table";
        }

        /// <summary>Combined rule-set fingerprint: sorted config_key=sha256 lines plus the evaluated goal-table definition hash.</summary>
        public static string RulesetSha256(IEnumerable<ReviewLogConfigurationInput> inputs, string goalTableSha256)
        {
            var lines = (inputs ?? new ReviewLogConfigurationInput[0]).Where(i => i != null)
                .Select(i => (i.ConfigKey ?? string.Empty) + "=" + (i.Sha256 ?? i.Availability ?? string.Empty).ToLowerInvariant())
                .OrderBy(l => l, StringComparer.Ordinal).ToList();
            if (!string.IsNullOrEmpty(goalTableSha256)) lines.Add("goal_table=" + goalTableSha256);
            return lines.Count == 0 ? string.Empty : ReviewLogRulesetResolver.Sha256Hex(string.Join("\n", lines));
        }

        private static string MappingStatus(ReviewSnapshot snapshot, ReviewPqmRow row)
        {
            ReviewStructureMapping mapping = snapshot.StructureMappings.FirstOrDefault(m => m != null &&
                string.Equals(m.TemplateStructure, row.TemplateStructure, StringComparison.Ordinal));
            string status = mapping == null ? null : mapping.Status;
            if (status == ReviewStatusCodes.Pass) return "resolved";
            if (status == ReviewStatusCodes.Variation) return "ambiguous";
            if (status != null) return "unresolved";
            return string.IsNullOrWhiteSpace(row.ResolvedStructureId) ? "unresolved" : "resolved";
        }

        private static string[] CheckRow(string id, int rowNo, ReviewCheckRow row, string status, bool included, Func<string, string> text)
        {
            bool numeric = !string.Equals(row.Unit, ReviewUnitCodes.Text, StringComparison.Ordinal) &&
                !string.Equals(row.Unit, ReviewUnitCodes.Boolean, StringComparison.Ordinal);
            double observed, expected;
            return new[]
            {
                id, rowNo.ToString(CultureInfo.InvariantCulture), row.CheckCode ?? string.Empty, row.Category ?? string.Empty,
                row.Severity ?? string.Empty, status, included ? "1" : "0",
                text(row.ObservedValue), text(row.ExpectedValue),
                numeric && ReviewLogCsv.TryParseInvariantNumber(row.ObservedValue, out observed) ? ReviewLogCsv.Number(observed) : string.Empty,
                numeric && ReviewLogCsv.TryParseInvariantNumber(row.ExpectedValue, out expected) ? ReviewLogCsv.Number(expected) : string.Empty,
                row.Unit ?? string.Empty, text(row.Message)
            };
        }

        private static void AddParameters(ReviewLogRecord record, string id, ReviewPlanAnalysis analysis, Func<string, string> text)
        {
            if (analysis == null) return;
            var rows = record.Rows[ReviewLogSchema.PlanParameterTable];
            Action<string, string, string, double?, string, string, string> add = (scope, entity, name, value, unit, method, note) =>
            {
                bool available = value.HasValue && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value);
                rows.Add(new[] { id, (rows.Count + 1).ToString(CultureInfo.InvariantCulture), scope, entity ?? string.Empty, name, ReviewLogCsv.Number(value), unit, method,
                    available ? "available" : "unavailable", text(note) });
            };
            string pamMethod = "pam_" + (analysis.PamWeightingMode ?? "unknown").ToLowerInvariant() + "_weighted";
            string smallMethod = "fraction_cp_area_lt_" + ReviewLogCsv.Number(analysis.SmallApertureThresholdCm2) + "cm2_mu_weighted";
            add("plan", null, "total_mu", analysis.TotalMetersetMu, "MU", "sum_beam_meterset", analysis.TotalMetersetMu.HasValue ? null : "Meterset unavailable.");
            add("plan", null, "mu_per_gy", analysis.MuPerGy, "MU/Gy", "total_mu_per_fraction_dose", null);
            add("plan", analysis.TargetStructureId, "pam", analysis.Pam, "1", pamMethod,
                analysis.Pam.HasValue ? analysis.TargetSelectionProvenance : analysis.PamReason);
            add("plan", null, "mean_aperture_cm2", analysis.MeanApertureAreaCm2, "cm2", "mu_weighted_mean_cp_aperture", null);
            add("plan", null, "small_aperture_fraction", analysis.SmallApertureFraction, "1", smallMethod, null);
            foreach (ReviewTargetQuality target in analysis.TargetQuality.Where(t => t != null))
            {
                string note = target.Note;
                add("target", target.StructureId, "paddick_ci", target.PaddickCi, "1", "paddick_tv_rx2_over_tv_vrx", note);
                add("target", target.StructureId, "paddick_ci_reciprocal", target.PlanCheckCi, "1", "inverse_paddick", null);
                add("target", target.StructureId, "gradient_index", target.GradientIndex, "1", "body_v50_over_v100", null);
                add("target", target.StructureId, "homogeneity_index", target.HomogeneityIndex, "1", "d2_minus_d98_over_rx", null);
                add("target", target.StructureId, "target_volume_cm3", target.TargetVolumeCm3, "cm3", "structure_volume", null);
                add("target", target.StructureId, "reference_dose_gy", target.ReferenceDoseGy, "Gy", "plan_prescription", null);
            }
            foreach (ReviewBeamAnalysis beam in analysis.Beams.Where(b => b != null))
            {
                string entity = beam.BeamNumber.HasValue ? "beam-" + beam.BeamNumber.Value.ToString(CultureInfo.InvariantCulture) : (beam.BeamId ?? string.Empty);
                add("beam", entity, "total_mu", beam.MetersetMu, "MU", "beam_meterset", beam.MetersetReason);
                add("beam", entity, "pam", beam.Pam, "1", pamMethod, beam.Pam.HasValue ? null : beam.PamReason);
                add("beam", entity, "mean_aperture_cm2", beam.MeanApertureAreaCm2, "cm2", "mu_weighted_mean_cp_aperture",
                    beam.MeanApertureAreaCm2.HasValue ? null : beam.GeometryReason);
                add("beam", entity, "small_aperture_fraction", beam.SmallApertureFraction, "1", smallMethod, null);
                add("beam", entity, "estimated_duration_s", beam.EstimatedBeamDurationSeconds, "s", "dose_rate_profile_estimate",
                    beam.EstimatedBeamDurationSeconds.HasValue ? beam.DoseRateEstimateProfile : beam.DoseRateEstimateReason);
            }
        }

        private static string Distinct(IEnumerable<string> values)
        {
            return string.Join("|", values.Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim())
                .Distinct(StringComparer.Ordinal).OrderBy(v => v, StringComparer.Ordinal));
        }
    }
}
