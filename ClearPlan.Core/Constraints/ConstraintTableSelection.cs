using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public sealed class ConstraintTableSelectionCandidate
    {
        public ConstraintTableSelectionCandidate()
        {
            Reasons = new List<string>();
        }

        public ConstraintTableDefinition Table { get; set; }
        public int Score { get; set; }
        public IList<string> Reasons { get; set; }
    }

    public sealed class ConstraintTableSelection
    {
        public ConstraintTableSelection()
        {
            Candidates = new List<ConstraintTableSelectionCandidate>();
        }

        public ConstraintTableDefinition SelectedTable { get; set; }
        public bool RequiresConfirmation { get; set; }
        public string Reason { get; set; }
        public IList<ConstraintTableSelectionCandidate> Candidates { get; set; }
    }
}
