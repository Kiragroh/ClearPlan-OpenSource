using System.Collections.Generic;

namespace ClearPlan.Core.Constraints
{
    public sealed class ConstraintCatalog
    {
        public ConstraintCatalog()
        {
            Structures = new List<StructureDefinition>();
            Tables = new List<ConstraintTableDefinition>();
            Issues = new List<CatalogValidationIssue>();
            Statistics = new CatalogStatistics();
        }

        public string Schema { get; set; }
        public string SourceKind { get; set; }
        public string SourcePath { get; set; }
        public IList<StructureDefinition> Structures { get; set; }
        public IList<ConstraintTableDefinition> Tables { get; set; }
        public IList<CatalogValidationIssue> Issues { get; set; }
        public CatalogStatistics Statistics { get; set; }
    }
}
