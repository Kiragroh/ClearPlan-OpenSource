using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using ClearPlan;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Core.ReviewLog;
using ClearPlan.Helpers;
using ClearPlan.Review;
using VMS.TPS.Common.Model.API;

// Read-only ESAPI standalone executable. It opens no modification session,
// writes no report, uploads nothing and changes no TPS/OIS state.
[assembly: ESAPIScript(IsWriteable = false)]

namespace ClearPlan.ReviewLogMining
{
    /// <summary>One explicit plan reference from the authorised input list.</summary>
    public sealed class MiningTarget
    {
        public int LineNumber { get; set; }
        public string PatientId { get; set; }
        public string CourseId { get; set; }
        public string PlanId { get; set; }
        public string PlanUid { get; set; }
    }

    public static class MiningInput
    {
        public static readonly string[] RequiredColumns = { "patient_id", "course_id", "plan_id" };

        /// <summary>RFC 4180 CSV with header patient_id,course_id,plan_id[,plan_uid]. Blank/# lines are ignored.</summary>
        public static List<MiningTarget> Parse(string csvText)
        {
            List<List<string>> rows = ReviewLogCsv.Parse(csvText ?? string.Empty)
                .Where(r => !(r.Count > 0 && r[0].TrimStart().StartsWith("#", StringComparison.Ordinal))).ToList();
            if (rows.Count == 0) throw new FormatException("The plan list is empty.");
            var header = rows[0].Select(h => h.Trim().ToLowerInvariant()).ToList();
            foreach (string required in RequiredColumns)
                if (!header.Contains(required)) throw new FormatException("The plan list requires a '" + required + "' column.");
            int patient = header.IndexOf("patient_id"), course = header.IndexOf("course_id"), plan = header.IndexOf("plan_id"), uid = header.IndexOf("plan_uid");
            var targets = new List<MiningTarget>();
            for (int i = 1; i < rows.Count; i++)
            {
                List<string> row = rows[i];
                Func<int, string> cell = index => index >= 0 && index < row.Count ? row[index].Trim() : string.Empty;
                var target = new MiningTarget { LineNumber = i + 1, PatientId = cell(patient), CourseId = cell(course), PlanId = cell(plan), PlanUid = cell(uid) };
                if (target.PatientId.Length == 0 || target.CourseId.Length == 0 || target.PlanId.Length == 0)
                    throw new FormatException("Plan list row " + target.LineNumber + " lacks patient_id, course_id or plan_id.");
                targets.Add(target);
            }
            return targets;
        }
    }

    public static class Program
    {
        private const string Usage =
            "ClearPlan.RetrospectiveReviewMining --input <plan-list.csv> [--log-dir <dir>] [--validate-input-only]\n" +
            "Read-only retrospective review mining. Writes review-log rows with mode = retrospective.\n" +
            "Rules: settings.ini/settings.json next to ClearPlan.esapi.dll. No report, no OIS upload, no TPS writes.";

        [STAThread]
        public static int Main(string[] args)
        {
            string input = null, logDirectory = null;
            bool validateOnly = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "--input" && i + 1 < args.Length) input = args[++i];
                else if (args[i] == "--log-dir" && i + 1 < args.Length) logDirectory = args[++i];
                else if (args[i] == "--validate-input-only") validateOnly = true;
                else { Console.Error.WriteLine(Usage); return 2; }
            }
            if (string.IsNullOrWhiteSpace(input)) { Console.Error.WriteLine(Usage); return 2; }

            List<MiningTarget> targets;
            try { targets = MiningInput.Parse(File.ReadAllText(input, Encoding.UTF8)); }
            catch (Exception error) { Console.Error.WriteLine("Plan list rejected: " + error.Message); return 2; }
            Console.WriteLine("Plan list accepted: " + targets.Count + " plan(s).");
            if (validateOnly) return 0;

            var settings = ClearPlanSettings.Load();
            if (!string.IsNullOrWhiteSpace(logDirectory)) settings.Paths.ReviewLogDirectory = Path.GetFullPath(logDirectory);
            // An explicit batch run is itself the authorisation to log; it still honours the privacy switches.
            settings.ReviewLog.Enabled = true;

            // WPF resources/dispatcher exist as in the Runner host; never shown.
            var wpf = System.Windows.Application.Current ?? new System.Windows.Application { ShutdownMode = System.Windows.ShutdownMode.OnExplicitShutdown };
            int ok = 0, partial = 0, failed = 0;
            using (VMS.TPS.Common.Model.API.Application app = VMS.TPS.Common.Model.API.Application.CreateApplication())
            {
                foreach (MiningTarget target in targets)
                {
                    string status = ProcessOne(app, target, settings);
                    if (status == ReviewLogRunStatus.Ok) ok++;
                    else if (status == ReviewLogRunStatus.Partial) partial++;
                    else failed++;
                    // Row number only: identifiers are never echoed to the console.
                    Console.WriteLine("row " + target.LineNumber + ": " + status);
                }
            }
            Console.WriteLine("Retrospective mining finished: ok=" + ok + " partial=" + partial + " failed=" + failed + ".");
            return failed == 0 ? 0 : 1;
        }

        internal static string ProcessOne(VMS.TPS.Common.Model.API.Application app, MiningTarget target, ClearPlanSettings settings)
        {
            DateTimeOffset started = DateTimeOffset.UtcNow;
            ReviewLogContext context = null;
            ReviewSnapshot snapshot = null;
            MainViewModel vm = null;
            var notes = new List<string>();
            string status = ReviewLogRunStatus.Ok;
            try
            {
                app.ClosePatient();
                Patient patient = app.OpenPatientById(target.PatientId);
                if (patient == null) return WriteFailure(target, settings, started, "patient-not-found");
                Course course = patient.Courses.FirstOrDefault(c => string.Equals(c.Id, target.CourseId, StringComparison.Ordinal));
                PlanSetup plan = course == null ? null : course.PlanSetups.FirstOrDefault(p => string.Equals(p.Id, target.PlanId, StringComparison.Ordinal));
                if (plan == null) return WriteFailure(target, settings, started, course == null ? "course-not-found" : "plan-not-found");
                if (target.PlanUid.Length > 0 && !string.Equals(plan.UID, target.PlanUid, StringComparison.Ordinal))
                    return WriteFailure(target, settings, started, "plan-uid-mismatch");

                // Same review path as the headless report entry: identical snapshot builder and rule set.
                var item = new PlanningItemViewModel(plan);
                vm = new MainViewModel(app.CurrentUser, patient, "retrospective review mining",
                    new ObservableCollection<PlanningItemViewModel> { item }, item, plan);
                ClearPlan.PhantomValidation.Selection.Apply(vm, patient, plan);
                vm.PqmSummaries = new ObservableCollection<PQMSummaryViewModel>();
                if (vm.ActiveConstraintPath != null && vm.ConstraintCatalogStatus != null && vm.ConstraintCatalogStatus.IsUsable)
                    vm.GetPQMSummaries(vm.ActiveConstraintPath, item, patient);
                vm.ApplyDefaultDvhSelections();
                snapshot = new EsapiReviewSnapshotBuilder().Build(vm, settings);
                if (vm.ConstraintSelectionRequiresConfirmation || vm.PqmSummaries.Count == 0)
                {
                    snapshot.PlanCheckRows.Insert(0, new ReviewCheckRow
                    {
                        CheckCode = "constraint-selection-unconfirmed", Category = "PQM", Status = ReviewStatusCodes.Variation,
                        Severity = ReviewSeverityCodes.Warning, Unit = ReviewUnitCodes.Text,
                        Message = "Constraint table selection requires confirmation; goal assessment remains unevaluated."
                    });
                }
                try
                {
                    NativeMlcProfileCatalog profiles = null;
                    DoseRateEstimationProfileCatalog rates = null;
                    try { profiles = NativeMlcProfileFileSource.LoadAsync(settings.ResolvePath(settings.Paths.MlcGeometryProfilesJsonPath), CancellationToken.None).GetAwaiter().GetResult(); }
                    catch (Exception) { notes.Add("mlc-profile-unavailable"); }
                    try { rates = DoseRateProfileFileSource.LoadAsync(string.IsNullOrWhiteSpace(settings.Paths.DoseRateProfilesJsonPath) ? null : settings.ResolvePath(settings.Paths.DoseRateProfilesJsonPath), CancellationToken.None).GetAwaiter().GetResult(); }
                    catch (Exception) { notes.Add("dose-rate-profile-unavailable"); }
                    var builder = new EsapiPlanAnalysisBuilder(profiles, rates);
                    ReviewPlanAnalysis analysis = builder.Build(plan);
                    using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(120)))
                        analysis = builder.EnrichTargetProjectionsAsync(plan, analysis, timeout.Token).GetAwaiter().GetResult();
                    snapshot.PlanAnalysis = analysis;
                }
                catch (Exception error)
                {
                    status = ReviewLogRunStatus.Partial;
                    notes.Add("plan-descriptors-unavailable:" + error.GetType().Name);
                }
                context = EsapiReviewLog.CreateContext(ReviewLogModes.Retrospective, app.CurrentUser, patient, plan, EsapiReviewLog.GoalTableId(vm), started);
                context.ReportWritten = false;
                context.RunStatus = status;
                context.RunMessage = string.Join("; ", notes);
                string failure;
                if (!EsapiReviewLog.TryWrite(snapshot, context, settings, EsapiReviewLog.ActiveConstraintSource(vm), 30000, out failure))
                {
                    Console.Error.WriteLine("row " + target.LineNumber + ": review log not written (" + failure + ")");
                    return ReviewLogRunStatus.Failed;
                }
                return status;
            }
            catch (Exception error)
            {
                // Type only: native messages can contain identifiers.
                return WriteFailure(target, settings, started, "review-failed:" + error.GetType().Name);
            }
            finally
            {
                try { app.ClosePatient(); } catch (Exception) { }
            }
        }

        private static string WriteFailure(MiningTarget target, ClearPlanSettings settings, DateTimeOffset started, string message)
        {
            ReviewLogContext context = EsapiReviewLog.CreateContext(ReviewLogModes.Retrospective, null, null, null, null, started);
            context.FinishedUtc = DateTimeOffset.UtcNow;
            context.PatientId = target.PatientId;
            context.CourseId = target.CourseId;
            context.PlanId = target.PlanId;
            context.PlanUid = target.PlanUid.Length == 0 ? null : target.PlanUid;
            context.PlanningItemType = "plan";
            context.RunStatus = ReviewLogRunStatus.Failed;
            context.RunMessage = message;
            context.SensitiveTerms.Add(target.PatientId);
            string failure;
            EsapiReviewLog.TryWrite(null, context, settings, null, 30000, out failure);
            return ReviewLogRunStatus.Failed;
        }
    }
}
