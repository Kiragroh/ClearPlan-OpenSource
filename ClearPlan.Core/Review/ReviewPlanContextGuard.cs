using System;

namespace ClearPlan.Core.Review
{
    /// <summary>Host-only identity guard; no UID is added to the serializable review contract.</summary>
    public sealed class ReviewPlanContextGuard
    {
        private ReviewSnapshot committedSnapshot;
        private string committedPlanUid;
        private object committedPlanningObject;

        public void Commit(ReviewSnapshot snapshot, string planUid, object planningObject = null)
        {
            if (snapshot == null || snapshot.Synthetic || (string.IsNullOrWhiteSpace(planUid) && planningObject == null))
                throw new ArgumentException("Only a complete clinical context can be committed.");
            committedSnapshot = snapshot;
            committedPlanUid = planUid;
            committedPlanningObject = planningObject;
        }

        public bool Matches(ReviewSnapshot snapshot, string livePlanUid, object planningObject = null)
        {
            return snapshot != null && !snapshot.Synthetic &&
                ReferenceEquals(snapshot, committedSnapshot) &&
                (committedPlanningObject == null || ReferenceEquals(committedPlanningObject, planningObject)) &&
                (string.IsNullOrWhiteSpace(committedPlanUid)
                    // PlanSum has no SOP UID in the host. Use only its owning
                    // session's object identity; never invent a DICOM identifier.
                    ? committedPlanningObject != null && string.IsNullOrWhiteSpace(livePlanUid)
                    : string.Equals(committedPlanUid, livePlanUid, StringComparison.Ordinal));
        }

        public void Clear() { committedSnapshot = null; committedPlanUid = null; committedPlanningObject = null; }
    }
}
