using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.ReviewLog
{
    /// <summary>
    /// Stable, append-only review-log contract (schema version 1). Column names are part of the
    /// public data dictionary (docs/REVIEW_LOG.md, tools/review-log/schema.sql); never rename or
    /// reorder them within a schema version. Any change requires a new schema version and new files.
    /// </summary>
    public static class ReviewLogSchema
    {
        public const int SchemaVersion = 1;

        public const string ReviewTable = "review";
        public const string GoalResultTable = "goal_result";
        public const string PlanParameterTable = "plan_parameter";
        public const string CheckFindingTable = "check_finding";
        public const string RulesetRevisionTable = "ruleset_revision";

        public static readonly string[] ReviewColumns =
        {
            "schema_version", "review_id", "mode", "started_utc", "finished_utc", "clearplan_version", "build_id",
            "identifier_mode", "user_key", "patient_key", "course_key", "plan_key", "plan_label", "planning_item_type",
            "plan_created_utc", "plan_approval_status", "planning_approval_utc", "treatment_approval_utc",
            "machine_id", "technique", "mlc_model", "total_dose_gy", "dose_per_fraction_gy", "fractions",
            "goal_table_id", "goal_table_sha256", "ruleset_sha256",
            "n_goals", "n_goal", "n_variation", "n_not_met", "n_not_evaluated",
            "n_checks", "n_checks_excluded", "n_findings_open",
            "report_written", "report_sha256", "report_filed", "run_status", "run_message"
        };

        public static readonly string[] GoalResultColumns =
        {
            "review_id", "row_no", "stable_id", "source_kind", "template_code", "template_structure",
            "resolved_structure", "mapping_status", "objective", "comparator", "goal_value", "variation_value",
            "achieved_value", "unit", "status", "priority", "source_label"
        };

        public static readonly string[] PlanParameterColumns =
        {
            "review_id", "row_no", "scope", "entity", "name", "value", "unit", "method", "availability", "note"
        };

        public static readonly string[] CheckFindingColumns =
        {
            "review_id", "row_no", "check_code", "category", "severity", "status", "included",
            "observed_value", "expected_value", "observed_numeric", "expected_numeric", "unit", "message"
        };

        public static readonly string[] RulesetRevisionColumns =
        {
            "review_id", "config_key", "category", "revision_number", "sha256", "revision_created_utc", "availability"
        };

        /// <summary>Write order: child tables first, the review row last (a review row implies complete children).</summary>
        public static readonly string[] WriteOrder =
        {
            RulesetRevisionTable, GoalResultTable, PlanParameterTable, CheckFindingTable, ReviewTable
        };

        public static string FileName(string table)
        {
            return table + ".csv";
        }

        public static IList<string> Columns(string table)
        {
            switch (table)
            {
                case ReviewTable: return ReviewColumns;
                case GoalResultTable: return GoalResultColumns;
                case PlanParameterTable: return PlanParameterColumns;
                case CheckFindingTable: return CheckFindingColumns;
                case RulesetRevisionTable: return RulesetRevisionColumns;
                default: throw new ArgumentOutOfRangeException("table", table, "Unknown review-log table.");
            }
        }
    }

    public static class ReviewLogModes
    {
        public const string Live = "live";
        public const string Headless = "headless";
        public const string Retrospective = "retrospective";

        public static bool IsKnown(string mode)
        {
            return mode == Live || mode == Headless || mode == Retrospective;
        }
    }

    public static class ReviewLogRunStatus
    {
        public const string Ok = "ok";
        public const string Partial = "partial";
        public const string Failed = "failed";
    }

    /// <summary>Goal-row status vocabulary. Anything that is not an explicit pass/variation/fail is not_evaluated.</summary>
    public static class ReviewLogGoalStatus
    {
        public const string Goal = "goal";
        public const string Variation = "variation";
        public const string NotMet = "not_met";
        public const string NotEvaluated = "not_evaluated";

        public static string FromSnapshot(string snapshotStatus)
        {
            switch ((snapshotStatus ?? string.Empty).Trim().ToLowerInvariant())
            {
                case ReviewStatusCodes.Pass: return Goal;
                case ReviewStatusCodes.Variation: return Variation;
                case ReviewStatusCodes.Fail: return NotMet;
                default: return NotEvaluated;
            }
        }
    }

    /// <summary>Check-row status vocabulary; OK rows are logged so that denominators exist.</summary>
    public static class ReviewLogCheckStatus
    {
        public const string Ok = "ok";
        public const string Variation = "variation";
        public const string Fail = "fail";
        public const string Info = "info";
        public const string NotEvaluated = "not_evaluated";

        public static string FromSnapshot(string snapshotStatus)
        {
            switch ((snapshotStatus ?? string.Empty).Trim().ToLowerInvariant())
            {
                case ReviewStatusCodes.Pass: return Ok;
                case ReviewStatusCodes.Variation: return Variation;
                case ReviewStatusCodes.Fail: return Fail;
                case ReviewStatusCodes.Info: return Info;
                default: return NotEvaluated;
            }
        }

        public static bool IsOpen(string logStatus)
        {
            return logStatus == Variation || logStatus == Fail || logStatus == NotEvaluated;
        }
    }
}
