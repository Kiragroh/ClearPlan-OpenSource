using System;
using System.Collections.Generic;
using ClearPlan.Core.Review;

namespace ClearPlan.Core.ReviewLog
{
    /// <summary>Review-log switches. Disabled by default; pseudonymised identifiers by default.</summary>
    public sealed class ReviewLogOptions
    {
        public bool Enabled { get; set; }
        public string Directory { get; set; }
        public string KeyFilePath { get; set; }
        /// <summary>Write clear patient/course/plan/user identifiers instead of HMAC pseudonyms (local policy only).</summary>
        public bool LogClearIdentifiers { get; set; }
        /// <summary>Write the plan ID as plan_label. Off by default: plan IDs can carry initials or names.</summary>
        public bool LogPlanLabel { get; set; }
        /// <summary>Write check messages/notes (scrubbed for known identifiers). On by default.</summary>
        public bool LogMessages { get; set; } = true;
        /// <summary>Maximum wait for the cross-process append lock.</summary>
        public int LockTimeoutMilliseconds { get; set; } = 5000;
    }

    /// <summary>One configuration document that defined the rules of a review.</summary>
    public sealed class ReviewLogConfigurationInput
    {
        public string ConfigKey { get; set; }
        public string Category { get; set; }
        public int? RevisionNumber { get; set; }
        public string Sha256 { get; set; }
        public DateTimeOffset? RevisionCreatedUtc { get; set; }
        /// <summary>available | unavailable | not_configured | not_hashed</summary>
        public string Availability { get; set; }
    }

    /// <summary>Per-review facts that are not part of the detached snapshot (identity, timing, provenance).</summary>
    public sealed class ReviewLogContext
    {
        public ReviewLogContext()
        {
            ReviewId = Guid.NewGuid();
            ConfigurationInputs = new List<ReviewLogConfigurationInput>();
            ExcludedCheckRows = new List<ReviewCheckRow>();
            SensitiveTerms = new List<string>();
            RunStatus = ReviewLogRunStatus.Ok;
        }

        public Guid ReviewId { get; set; }
        public string Mode { get; set; }
        public DateTimeOffset? StartedUtc { get; set; }
        public DateTimeOffset? FinishedUtc { get; set; }
        public string ClearPlanVersion { get; set; }
        public string BuildId { get; set; }
        // Raw identifiers stay in memory only; they are pseudonymised (or written clear on explicit policy) by the builder.
        public string UserId { get; set; }
        public string PatientId { get; set; }
        public string CourseId { get; set; }
        public string PlanUid { get; set; }
        public string PlanId { get; set; }
        public string PlanningItemType { get; set; }
        public DateTimeOffset? PlanCreatedUtc { get; set; }
        public string PlanApprovalStatus { get; set; }
        public DateTimeOffset? PlanningApprovalUtc { get; set; }
        public DateTimeOffset? TreatmentApprovalUtc { get; set; }
        public string GoalTableId { get; set; }
        public bool ReportWritten { get; set; }
        public string ReportSha256 { get; set; }
        public bool ReportFiled { get; set; }
        public string RunStatus { get; set; }
        public string RunMessage { get; set; }
        public List<ReviewLogConfigurationInput> ConfigurationInputs { get; set; }
        /// <summary>Checks calculated but excluded by the configured check selection (logged with included = 0).</summary>
        public List<ReviewCheckRow> ExcludedCheckRows { get; set; }
        /// <summary>Patient/user names and IDs to be redacted from any free-text column.</summary>
        public List<string> SensitiveTerms { get; set; }
    }
}
