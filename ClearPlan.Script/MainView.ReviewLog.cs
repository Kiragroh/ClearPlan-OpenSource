using System;
using ClearPlan.Core.Review;
using ClearPlan.Core.ReviewLog;
using ClearPlan.Review;

namespace ClearPlan
{
    /// <summary>
    /// Live review-log hook (mode = live). One row set per clinical snapshot: written when a plan report is
    /// saved (report_written = 1) or, if no report was saved for the current snapshot, when the review closes.
    /// Logging can never fail the review: every path is guarded and bounded in time.
    /// The legacy LiveMining.csv writer (LogLiveMining) is left unchanged.
    /// </summary>
    public partial class MainView
    {
        private const int ReviewLogMaxWaitMilliseconds = 3000;
        private readonly DateTimeOffset _reviewLogStartedUtc = DateTimeOffset.UtcNow;
        private ReviewSnapshot _reviewLogLastSnapshot;

        private void WriteLiveReviewLog(ReviewSnapshot snapshot, string reportPath)
        {
            try
            {
                if (snapshot == null || snapshot.Synthetic || ReferenceEquals(snapshot, _reviewLogLastSnapshot)) return;
                if (!_settings.ReviewLog.Enabled) return;
                var item = _vm.ActivePlanningItem == null ? null : _vm.ActivePlanningItem.PlanningItemObject;
                ReviewLogContext context = EsapiReviewLog.CreateContext(ReviewLogModes.Live, _vm.User, _vm.Patient, item,
                    EsapiReviewLog.GoalTableId(_vm), _reviewLogStartedUtc);
                context.ReportWritten = reportPath != null;
                context.ReportSha256 = reportPath == null ? null : EsapiReviewLog.FileSha256OrNull(reportPath);
                string failure;
                if (EsapiReviewLog.TryWrite(snapshot, context, _settings, EsapiReviewLog.ActiveConstraintSource(_vm),
                    ReviewLogMaxWaitMilliseconds, out failure))
                    _reviewLogLastSnapshot = snapshot;
                else if (failure != "disabled")
                    log.Warn("Structured review log not written: " + failure);
            }
            catch (Exception error)
            {
                try { log.Warn("Structured review log not written: " + error.GetType().Name); } catch (Exception) { }
            }
        }

        private void WriteLiveReviewLogOnClose()
        {
            try
            {
                if (_clinicalReviewHost == null) return;
                WriteLiveReviewLog(_clinicalReviewHost.CurrentClinicalSnapshotForLog, null);
            }
            catch (Exception) { }
        }
    }
}
