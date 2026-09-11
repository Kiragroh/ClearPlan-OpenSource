using System;
using System.Collections.Generic;
using System.Globalization;
using ClearPlan.Core.Review;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed class ReviewPqmRowViewModel
    {
        public ReviewPqmRowViewModel(ReviewPqmRow row)
        {
            if (row == null)
            {
                throw new ArgumentNullException("row");
            }

            StableId = row.StableId;
            TemplateCode = row.TemplateCode;
            TemplateStructure = row.TemplateStructure;
            ResolvedStructureId = row.ResolvedStructureId;
            StructureOptions = new List<string>(
                row.StructureOptions ?? new List<string>());
            Objective = row.Objective;
            Comparator = row.Comparator;
            Goal = FormatValue(row.Goal);
            Variation = FormatValue(row.Variation);
            AchievedValue = FormatValue(row.AchievedValue);
            Unit = row.Unit;
            StatusCode = row.Status;
            StatusText = ReviewDisplayText.Status(row.Status);
            SeverityCode = row.Severity;
            SeverityText = ReviewDisplayText.Severity(row.Severity);
            Explanation = row.Explanation;
            SourceLabel = row.SourceLabel;
            MappingDescription = row.MappingDescription;
        }

        public string StableId { get; private set; }

        public string TemplateCode { get; private set; }

        public string TemplateStructure { get; private set; }

        public string ResolvedStructureId { get; private set; }

        public IList<string> StructureOptions { get; private set; }

        public string Objective { get; private set; }

        public string Comparator { get; private set; }

        public string Goal { get; private set; }

        public string Variation { get; private set; }

        public string AchievedValue { get; private set; }

        public string Unit { get; private set; }

        public string StatusCode { get; private set; }

        public string StatusText { get; private set; }

        public string SeverityCode { get; private set; }

        public string SeverityText { get; private set; }

        public string Explanation { get; private set; }
        public string SourceLabel { get; private set; }
        public string MappingDescription { get; private set; }

        public string GoalText { get { return (string.IsNullOrWhiteSpace(Comparator) ? string.Empty : Comparator + " ") + WithUnit(Goal); } }
        public string VariationText { get { return WithUnit(Variation); } }
        public string AchievedText { get { return WithUnit(AchievedValue); } }
        public string StructureText { get { return string.IsNullOrWhiteSpace(ResolvedStructureId) ? "Nicht zugeordnet" : ResolvedStructureId; } }
        public string DetailsText { get { return TemplateCode + " · " + SeverityText + "\n" + MappingDescription + "\n" + SourceLabel + "\n" + Explanation; } }

        private string WithUnit(string value)
        { return value == "—" || string.IsNullOrWhiteSpace(Unit) ? value : value + " " + Unit; }

        private static string FormatValue(double? value)
        {
            return value.HasValue
                ? value.Value.ToString("0.###", CultureInfo.InvariantCulture)
                : "—";
        }
    }
}
