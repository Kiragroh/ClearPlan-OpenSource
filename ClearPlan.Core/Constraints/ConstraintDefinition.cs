using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public enum ConstraintDirection
    {
        Unknown,
        LowerIsBetter,
        HigherIsBetter
    }

    public sealed class ConstraintDefinition
    {
        public ConstraintDefinition()
        {
            Aliases = new List<string>();
            Codes = new List<string>();
            DicomTypes = new List<string>();
            RawValues = new Dictionary<string, string>();
        }

        public string ConstraintId { get; set; }
        public string TableId { get; set; }
        public string StructureId { get; set; }
        public string StructureName { get; set; }
        public string RawStructureName { get; set; }
        public string Metric { get; set; }
        public string Unit { get; set; }
        public string Comparator { get; set; }
        public ConstraintDirection ExpectedDirection { get; set; }
        public decimal? Goal { get; set; }
        public decimal? Variation { get; set; }
        public int? Priority { get; set; }
        public string Source { get; set; }
        public string Comment { get; set; }
        public string DvhObjective { get; set; }
        public string EvaluationPoint { get; set; }
        public IList<string> Aliases { get; set; }
        public IList<string> Codes { get; set; }
        public IList<string> DicomTypes { get; set; }
        public IDictionary<string, string> RawValues { get; set; }
    }
}
