using System;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan
{
    public sealed class ConstraintSourceStatusViewModel : ViewModelBase
    {
        public string Health { get; set; }
        public string Summary { get; set; }
        public string ActiveSource { get; set; }
        public string ActivePath { get; set; }
        public string LoadedAt { get; set; }
        public int Tables { get; set; }
        public int Constraints { get; set; }
        public int Structures { get; set; }
        public string Messages { get; set; }

        public static ConstraintSourceStatusViewModel From(
            ConstraintCatalogLoadResult result,
            bool refDbConfigured)
        {
            if (result == null)
            {
                return new ConstraintSourceStatusViewModel
                {
                    Health = "Error",
                    Summary = "Constraint-Quelle wurde noch nicht geladen.",
                    Messages = string.Empty
                };
            }

            bool warning = result.Warnings.Count > 0;
            string summary;
            if (!result.IsUsable)
            {
                summary = "Constraint-Quelle nicht verfügbar";
            }
            else if (warning &&
                     refDbConfigured &&
                     string.Equals(
                         result.ActiveSource,
                         "Excel",
                         StringComparison.OrdinalIgnoreCase))
            {
                summary = "RefDB nicht verfügbar – Excel aktiv";
            }
            else
            {
                summary = result.ActiveSource + " aktiv";
            }

            return new ConstraintSourceStatusViewModel
            {
                Health = !result.IsUsable ? "Error" : warning ? "Warning" : "Healthy",
                Summary = summary,
                ActiveSource = result.ActiveSource,
                ActivePath = result.ActivePath,
                LoadedAt = result.LoadedAtUtc.ToLocalTime().ToString("dd.MM.yyyy HH:mm:ss"),
                Tables = result.Catalog == null
                    ? 0
                    : result.Catalog.Statistics.ActiveTables,
                Constraints = result.Catalog == null
                    ? 0
                    : result.Catalog.Statistics.TotalConstraints,
                Structures = result.Catalog == null
                    ? 0
                    : result.Catalog.Statistics.ActiveStructures,
                Messages = string.Join(
                    Environment.NewLine,
                    result.Errors.Concat(result.Warnings))
            };
        }
    }
}
