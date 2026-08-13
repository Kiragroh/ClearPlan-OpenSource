using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public sealed class PlanConstraintContext
    {
        public PlanConstraintContext()
        {
            StructureIds = new List<string>();
        }

        public bool IsPlanSum { get; set; }
        public int? FractionCount { get; set; }
        public decimal? DosePerFractionGy { get; set; }
        public decimal? TotalDoseGy { get; set; }
        public string SiteHint { get; set; }
        public string RegimeHint { get; set; }
        public IList<string> StructureIds { get; set; }
    }
}
