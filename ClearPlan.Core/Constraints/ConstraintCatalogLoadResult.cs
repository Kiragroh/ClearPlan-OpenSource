using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.Constraints
{
    public sealed class ConstraintCatalogLoadResult
    {
        public ConstraintCatalogLoadResult()
        {
            ActiveSource = string.Empty;
            ActivePath = string.Empty;
            Warnings = new List<string>();
            Errors = new List<string>();
            LoadedAtUtc = DateTime.UtcNow;
        }

        public string ActiveSource { get; set; }
        public string ActivePath { get; set; }
        public ConstraintCatalog Catalog { get; set; }
        public IList<string> Warnings { get; set; }
        public IList<string> Errors { get; set; }
        public DateTime LoadedAtUtc { get; set; }

        public bool IsUsable
        {
            get
            {
                return Catalog != null &&
                       Catalog.Tables.Count > 0 &&
                       !Catalog.Issues.Any(issue => issue.IsFatal) &&
                       Errors.Count == 0;
            }
        }
    }
}
