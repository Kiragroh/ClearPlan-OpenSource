using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public enum StructureMatchRule
    {
        None,
        ExactRequestedName,
        CanonicalName,
        Alias,
        SideAlias,
        Code,
        DicomType
    }

    public sealed class StructureCandidate
    {
        public StructureCandidate()
        {
            Codes = new List<string>();
        }

        public string Id { get; set; }
        public string DicomType { get; set; }
        public IList<string> Codes { get; set; }
    }

    public sealed class StructureMatch
    {
        public StructureMatch()
        {
            CandidateIds = new List<string>();
        }

        public string CandidateId { get; set; }
        public StructureMatchRule Rule { get; set; }
        public int Confidence { get; set; }
        public bool IsAmbiguous { get; set; }
        public string Message { get; set; }
        public IList<string> CandidateIds { get; set; }
    }
}
