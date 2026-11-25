using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using ClearPlan.Core.Review;
using ClearPlan.Core.ReviewLog;
using ClearPlan.Helpers;
using VMS.TPS.Common.Model.API;

namespace ClearPlan.Review
{
    /// <summary>
    /// Native glue for the structured review log. Reads ESAPI properties only (no modification session,
    /// nothing is saved to the TPS); all formatting, pseudonymisation and file IO live in ClearPlan.Core.ReviewLog.
    /// Every public entry is guarded: logging failures are returned, never thrown into the review.
    /// </summary>
    public static class EsapiReviewLog
    {
        public static ReviewLogOptions Options(ClearPlanSettings settings)
        {
            var source = settings == null || settings.ReviewLog == null ? new ClearPlan.Core.Settings.ClearPlanReviewLogOptions() : settings.ReviewLog;
            return new ReviewLogOptions
            {
                Enabled = source.Enabled,
                LogClearIdentifiers = source.LogClearIdentifiers,
                LogPlanLabel = source.LogPlanLabel,
                LogMessages = source.LogMessages,
                Directory = settings == null ? null : settings.ResolvePath(settings.Paths.ReviewLogDirectory),
                KeyFilePath = settings == null ? null : settings.ResolvePath(settings.Paths.ReviewLogKeyFile)
            };
        }

        /// <summary>Must run on the ESAPI owner thread. Reads identity/approval metadata into a detached context.</summary>
        public static ReviewLogContext CreateContext(string mode, User user, Patient patient, PlanningItem item,
            string goalTableId, DateTimeOffset startedUtc)
        {
            var context = new ReviewLogContext
            {
                Mode = mode,
                StartedUtc = startedUtc,
                ClearPlanVersion = Safe(() => FileVersionInfo.GetVersionInfo(typeof(EsapiReviewLog).Assembly.Location).FileVersion),
                BuildId = Safe(() => typeof(EsapiReviewLog).Assembly.ManifestModule.ModuleVersionId.ToString("D")),
                UserId = Safe(() => user == null ? null : user.Id),
                PatientId = Safe(() => patient == null ? null : patient.Id),
                GoalTableId = goalTableId
            };
            AddTerms(context, () => patient == null ? null : patient.Id, () => patient == null ? null : patient.FirstName,
                () => patient == null ? null : patient.LastName, () => patient == null ? null : patient.MiddleName,
                () => user == null ? null : user.Id, () => user == null ? null : user.Name);
            if (user != null) foreach (string part in (Safe(() => user.Name) ?? string.Empty).Split(new[] { ' ', ',', '\\' }, StringSplitOptions.RemoveEmptyEntries)) context.SensitiveTerms.Add(part);
            var plan = item as PlanSetup;
            var sum = item as PlanSum;
            if (plan != null)
            {
                context.PlanningItemType = "plan";
                context.CourseId = Safe(() => plan.Course.Id);
                context.PlanUid = Safe(() => plan.UID);
                context.PlanId = Safe(() => plan.Id);
                context.PlanApprovalStatus = Safe(() => plan.ApprovalStatus.ToString());
                context.PlanCreatedUtc = ToUtc(SafeValue(() => plan.CreationDateTime));
                context.PlanningApprovalUtc = ParseLocal(Safe(() => plan.PlanningApprovalDate));
                context.TreatmentApprovalUtc = ParseLocal(Safe(() => plan.TreatmentApprovalDate));
                AddTerms(context, () => plan.PlanningApprover, () => plan.TreatmentApprover,
                    () => plan.PlanningApproverDisplayName, () => plan.TreatmentApproverDisplayName);
            }
            else if (sum != null)
            {
                context.PlanningItemType = "plan_sum";
                context.CourseId = Safe(() => sum.Course.Id);
                context.PlanId = Safe(() => sum.Id);
                context.PlanCreatedUtc = ToUtc(SafeValue(() => sum.CreationDateTime));
            }
            return context;
        }

        /// <summary>Hashes the configuration documents that defined this review (off the ESAPI thread is fine).</summary>
        public static List<ReviewLogConfigurationInput> DescribeRuleset(ClearPlanSettings settings, string activeConstraintSource)
        {
            var inputs = new List<ReviewLogConfigurationInput>();
            if (settings == null) return inputs;
            string managed = Safe(() => settings.ResolvePath(settings.Paths.ConfigurationDirectory));
            Func<string, string> resolve = path => string.IsNullOrWhiteSpace(path) ? null : Safe(() => settings.ResolvePath(path));
            string source = (activeConstraintSource ?? string.Empty).ToLowerInvariant();
            var mode = settings.ConstraintSource == null ? ClearPlan.Core.Constraints.ConstraintSourceMode.Automatic : settings.ConstraintSource.Mode;
            bool refDb = source.Contains("refdb") || (source.Length == 0 && mode == ClearPlan.Core.Constraints.ConstraintSourceMode.RefDb);
            bool excel = source.Contains("excel") || (source.Length == 0 && mode != ClearPlan.Core.Constraints.ConstraintSourceMode.RefDb);
            if (refDb) inputs.Add(ReviewLogRulesetResolver.Describe("constraints-refdb", "constraint_catalog", resolve(settings.ConstraintSource.RefDbJsonPath), managed));
            if (excel) inputs.Add(ReviewLogRulesetResolver.Describe("constraints-excel", "constraint_catalog", resolve(settings.ConstraintSource.ExcelWorkbookPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("structure-aliases", "structure_aliases", resolve(settings.ConstraintSource.StructureAliasesJsonPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("default-review-rules", "default_rules", resolve(settings.Paths.DefaultReviewRulesJsonPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("plancheck-selection", "check_selection", resolve(settings.Paths.PlanCheckSelectionJsonPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("ct-compatibility", "ct_approval", resolve(settings.Paths.CtCompatibilityJsonPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("field-naming", "field_naming", resolve(settings.Paths.FieldNamingRulesJsonPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("mlc-geometry-profiles", "machine_geometry", resolve(settings.Paths.MlcGeometryProfilesJsonPath), managed));
            inputs.Add(ReviewLogRulesetResolver.Describe("plancheck-resources", "legacy_plancheck", resolve(settings.ClinicalPaths.PlanCheckResourcesPath), managed));
            return inputs;
        }

        /// <summary>
        /// Writes one review. Returns false (with a short technical reason, no identifiers) when disabled or failed.
        /// The IO runs on a worker task and the caller waits at most <paramref name="maxWaitMilliseconds"/>,
        /// so an unreachable log share cannot freeze the review window indefinitely.
        /// </summary>
        public static bool TryWrite(ReviewSnapshot snapshot, ReviewLogContext context, ClearPlanSettings settings,
            string activeConstraintSource, int maxWaitMilliseconds, out string failure)
        {
            failure = null;
            try
            {
                ReviewLogOptions options = Options(settings);
                if (!options.Enabled) { failure = "disabled"; return false; }
                if (context == null) { failure = "no-context"; return false; }
                if (snapshot != null && snapshot.Synthetic) { failure = "synthetic-snapshot-not-logged"; return false; }
                if (snapshot != null && snapshot.ExcludedPlanCheckRows != null && context.ExcludedCheckRows.Count == 0)
                    context.ExcludedCheckRows.AddRange(snapshot.ExcludedPlanCheckRows);
                if (!context.FinishedUtc.HasValue) context.FinishedUtc = DateTimeOffset.UtcNow;
                Task<ReviewLogWriteResult> work = Task.Run(() =>
                {
                    context.ConfigurationInputs = DescribeRuleset(settings, activeConstraintSource);
                    return ReviewLogService.Write(snapshot, context, options);
                });
                if (!work.Wait(Math.Max(100, maxWaitMilliseconds)))
                {
                    work.ContinueWith(t => { var ignored = t.Exception; }, TaskContinuationOptions.OnlyOnFaulted);
                    failure = "timeout";
                    return false;
                }
                return true;
            }
            catch (AggregateException error)
            {
                failure = error.GetBaseException().GetType().Name;
                return false;
            }
            catch (Exception error)
            {
                // Type only: exception messages can contain paths or identifiers.
                failure = error.GetType().Name;
                return false;
            }
        }

        public static string GoalTableId(MainViewModel vm)
        {
            return Safe(() => vm == null || vm.ActiveConstraintPath == null ? null : vm.ActiveConstraintPath.ConstraintName);
        }

        public static string ActiveConstraintSource(MainViewModel vm)
        {
            return Safe(() => vm == null || vm.ConstraintCatalogStatus == null ? null : vm.ConstraintCatalogStatus.ActiveSource);
        }

        public static string FileSha256OrNull(string path)
        {
            return Safe(() => string.IsNullOrWhiteSpace(path) || !System.IO.File.Exists(path) ? null : ReviewLogRulesetResolver.FileSha256(path));
        }

        private static void AddTerms(ReviewLogContext context, params Func<string>[] readers)
        {
            foreach (Func<string> read in readers)
            {
                string value = Safe(read);
                if (!string.IsNullOrWhiteSpace(value)) context.SensitiveTerms.Add(value.Trim());
            }
        }

        private static DateTimeOffset? ToUtc(DateTime? local)
        {
            if (!local.HasValue) return null;
            DateTime value = local.Value.Kind == DateTimeKind.Unspecified ? DateTime.SpecifyKind(local.Value, DateTimeKind.Local) : local.Value;
            return new DateTimeOffset(value.ToUniversalTime(), TimeSpan.Zero);
        }

        /// <summary>ESAPI approval dates are local-time strings; unparsable values stay unavailable.</summary>
        private static DateTimeOffset? ParseLocal(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return null;
            DateTime parsed;
            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.AssumeLocal, out parsed) ||
                DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.AssumeLocal, out parsed))
                return new DateTimeOffset(parsed.ToUniversalTime(), TimeSpan.Zero);
            return null;
        }

        private static string Safe(Func<string> read)
        {
            try { return read(); } catch (Exception) { return null; }
        }

        private static T? SafeValue<T>(Func<T?> read) where T : struct
        {
            try { return read(); } catch (Exception) { return null; }
        }
    }
}
