using System.Collections.Generic;
using System.Linq;

namespace ClearPlan
{
    public sealed class OverviewViewModel : ViewModelBase
    {
        public int PqmPassed { get; set; }
        public int PqmAttention { get; set; }
        public int PqmFailed { get; set; }
        public int PlanCheckPassed { get; set; }
        public int PlanCheckWarnings { get; set; }
        public int PlanCheckErrors { get; set; }
        public int FieldNamesConforming { get; set; }
        public int FieldNamesDeviating { get; set; }
        public IList<string> TopDeviations { get; set; }

        public static OverviewViewModel Create(MainViewModel main)
        {
            IEnumerable<PQMSummaryViewModel> pqm =
                main.PqmSummaries ?? Enumerable.Empty<PQMSummaryViewModel>();
            IEnumerable<ErrorViewModel> checks =
                main.ErrorGrid ?? Enumerable.Empty<ErrorViewModel>();
            IEnumerable<FieldNamePreviewViewModel> fields =
                main.FieldNamePreviews ??
                Enumerable.Empty<FieldNamePreviewViewModel>();
            var deviations = pqm
                .Where(item => item.Met == "Not met" ||
                               item.Met == "Variation")
                .Select(item => "PQM · " + item.TemplateId + " · " +
                                item.DVHObjective + " · " + item.Met)
                .Concat(checks
                    .Where(item => item.Status != "3 - OK")
                    .Select(item => "PlanCheck · " + item.Description + " · " +
                                    item.Status))
                .Concat(fields
                    .Where(item => item.WouldChange)
                    .Select(item => "Feldname · " + item.CurrentId + " → " +
                                    item.SuggestedName))
                .Take(8)
                .ToList();
            if (deviations.Count == 0)
            {
                deviations.Add("Keine priorisierten Abweichungen.");
            }

            return new OverviewViewModel
            {
                PqmPassed = pqm.Count(item => item.Met == "Goal"),
                PqmAttention = pqm.Count(item => item.Met == "Variation" ||
                                                 item.Met == "Not evaluated"),
                PqmFailed = pqm.Count(item => item.Met == "Not met"),
                PlanCheckPassed = checks.Count(item => item.Status == "3 - OK"),
                PlanCheckWarnings = checks.Count(item =>
                    item.Status == "2 - Variation"),
                PlanCheckErrors = checks.Count(item =>
                    item.Status == "1 - Deviation"),
                FieldNamesConforming = fields.Count(item => !item.WouldChange),
                FieldNamesDeviating = fields.Count(item => item.WouldChange),
                TopDeviations = deviations
            };
        }
    }
}
