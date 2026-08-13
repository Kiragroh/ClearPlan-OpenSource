using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public sealed class StructureDefinition
    {
        public StructureDefinition()
        {
            Aliases = new List<string>();
            SideAliasesLeft = new List<string>();
            SideAliasesRight = new List<string>();
            DicomTypes = new List<string>();
            Codes = new List<string>();
        }

        public string StructureId { get; set; }
        public string CanonicalName { get; set; }
        public bool Active { get; set; }
        public string Laterality { get; set; }
        public IList<string> Aliases { get; set; }
        public IList<string> SideAliasesLeft { get; set; }
        public IList<string> SideAliasesRight { get; set; }
        public IList<string> DicomTypes { get; set; }
        public IList<string> Codes { get; set; }
    }
}
