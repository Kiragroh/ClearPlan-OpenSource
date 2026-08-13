using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public sealed class PqmObjectiveDefinition
    {
        public PqmObjectiveDefinition()
        {
            Aliases = new List<string>();
            Codes = new List<string>();
            DicomTypes = new List<string>();
        }

        public string ConstraintId { get; set; }
        public string TemplateId { get; set; }
        public IList<string> Aliases { get; set; }
        public IList<string> Codes { get; set; }
        public IList<string> DicomTypes { get; set; }
        public string DvhObjective { get; set; }
        public string Goal { get; set; }
        public string Variation { get; set; }
        public string Priority { get; set; }
        public string Source { get; set; }
        public string Comment { get; set; }
    }
}
