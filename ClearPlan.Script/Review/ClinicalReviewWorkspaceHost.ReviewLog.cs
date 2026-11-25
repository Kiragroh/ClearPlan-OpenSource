namespace ClearPlan.Review
{
    internal sealed partial class ClinicalReviewWorkspaceHost
    {
        /// <summary>The current clinical snapshot when it still matches the active plan; null in demo/stale/disposed state.</summary>
        internal ClearPlan.Core.Review.ReviewSnapshot CurrentClinicalSnapshotForLog
        {
            get
            {
                var snapshot = disposed ? null : controller.CurrentSnapshot;
                return snapshot != null && !snapshot.Synthetic && CanExportCurrentSnapshot(snapshot) ? snapshot : null;
            }
        }
    }
}
