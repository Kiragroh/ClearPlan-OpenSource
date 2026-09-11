using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public sealed class ConstraintTableDefinition
    {
        public ConstraintTableDefinition()
        {
            Constraints = new List<ConstraintDefinition>();
            PrescriptionLabels = new List<string>();
        }

        public string TableId { get; set; }
        public string DisplayName { get; set; }
        public bool Active { get; set; }
        public bool IsPlanSum { get; set; }
        public bool RequiresConfirmation { get; set; }
        public int? FractionCountMinimum { get; set; }
        public int? FractionCountMaximum { get; set; }
        public decimal? DosePerFractionMinimumGy { get; set; }
        public decimal? DosePerFractionMaximumGy { get; set; }
        public decimal? TotalDoseMinimumGy { get; set; }
        public decimal? TotalDoseMaximumGy { get; set; }
        public string Site { get; set; }
        public string Regime { get; set; }
        public string Source { get; set; }
        public IList<string> PrescriptionLabels { get; set; }
        public IList<ConstraintDefinition> Constraints { get; set; }
    }
}
