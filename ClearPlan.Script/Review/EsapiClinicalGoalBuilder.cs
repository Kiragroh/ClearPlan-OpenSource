using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using ClearPlan.Core.Review;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>
    /// Reads assigned Eclipse clinical goals on the owning ESAPI thread.
    /// Native evaluations are projected, not recalculated from RefDB constraints.
    /// </summary>
    public sealed class EsapiClinicalGoalBuilder
    {
        public const string SourceStableId = "source-eclipse-goals";

        private static readonly Regex ScalarWithUnit = new Regex(
            @"^\s*(?:<=|>=|<|>|=|\u2264|\u2265)?\s*(?<value>[+-]?(?:\d+(?:[.,]\d+)?|[.,]\d+)(?:[eE][+-]?\d+)?)\s*(?<unit>cGy|Gy|%|cc|cm3|cm\u00b3)\s*$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public List<ReviewPqmRow> Build(
            PlanningItem planningItem,
            out ReviewSourceStatus source)
        {
            if (System.Windows.Application.Current != null)
            {
                System.Windows.Application.Current.Dispatcher.VerifyAccess();
            }

            source = CreateSource(ReviewStatusCodes.Unavailable,
                "Eclipse clinical goals could not be read; no successful evaluation is inferred.");
            if (planningItem == null)
            {
                return new List<ReviewPqmRow>();
            }

            try
            {
                List<ClinicalGoal> goals = planningItem.GetClinicalGoals();
                List<string> structures = planningItem.StructureSet == null
                    ? new List<string>()
                    : planningItem.StructureSet.Structures
                        .Where(item => item != null && !item.IsEmpty)
                        .Select(item => item.Id).ToList();
                return MapGoals(goals, structures, out source);
            }
            catch (Exception)
            {
                // Optional source failure must not hide the other review sources.
                // Do not copy exception text or patient/session objects into DTOs.
                return new List<ReviewPqmRow>();
            }
        }

        /// <summary>
        /// Maps ESAPI value types and detached structure IDs, without accessing a session.
        /// </summary>
        public static List<ReviewPqmRow> MapGoals(
            IEnumerable<ClinicalGoal> goals,
            IEnumerable<string> availableStructureIds,
            out ReviewSourceStatus source)
        {
            var rows = new List<ReviewPqmRow>();
            if (goals == null)
            {
                source = CreateSource(ReviewStatusCodes.NotEvaluated,
                    "No assigned Eclipse Clinical Goals were returned by the API (null). RefDB goals remain a separate source; no native goal result is inferred.");
                return rows;
            }

            var available = new HashSet<string>(
                availableStructureIds ?? Enumerable.Empty<string>(),
                StringComparer.Ordinal);
            var usedIds = new HashSet<string>(StringComparer.Ordinal);
            foreach (ClinicalGoal goal in goals)
            {
                string structure = Safe(goal.StructureId, "Unassigned structure");
                bool matched = !string.IsNullOrWhiteSpace(goal.StructureId) &&
                    available.Contains(goal.StructureId) && structure == goal.StructureId;
                string objective = Safe(goal.ObjectiveAsString, goal.MeasureType.ToString());
                string comparator = MapComparator(goal.Objective.Operator);
                Match operatorMatch = Regex.Match(objective, @"<=|>=|<|>|=|\u2264|\u2265");
                string goalText = operatorMatch.Success
                    ? objective.Substring(operatorMatch.Index + operatorMatch.Length)
                    : string.Empty;
                string metric = operatorMatch.Success
                    ? objective.Substring(0, operatorMatch.Index).Trim()
                    : objective;

                string actualUnit;
                string goalUnit;
                string variationUnit;
                double? actual = ParseScalar(goal.ActualValueAsString, out actualUnit);
                double? limit = ParseScalar(goalText, out goalUnit);
                double? variation = ParseScalar(goal.VariationAcceptableAsString, out variationUnit);
                // Relative is explicitly documented as percentage. Absolute does
                // not distinguish Gy/cGy/volume, so it must never imply Gy here.
                if (!limit.HasValue && goal.Objective.LimitUnit == ObjectiveUnit.Relative)
                {
                    limit = Finite(goal.Objective.Limit);
                    goalUnit = ReviewUnitCodes.Percent;
                }
                string unit = actual.HasValue ? actualUnit : limit.HasValue
                    ? goalUnit : variation.HasValue ? variationUnit : ReviewUnitCodes.Text;
                if (actualUnit != unit) actual = null;
                if (goalUnit != unit) limit = null;
                if (variationUnit != unit) variation = null;

                ReviewStatusAndSeverity status = MapStatus(goal.EvaluationResult);
                if (!matched || status.Status == ReviewStatusCodes.NotEvaluated)
                {
                    status = new ReviewStatusAndSeverity(
                        ReviewStatusCodes.NotEvaluated, ReviewSeverityCodes.Info);
                    actual = null;
                }
                bool incomplete = (!limit.HasValue && goal.Objective.Operator != ObjectiveOperator.None) ||
                    (!actual.HasValue && status.Status != ReviewStatusCodes.NotEvaluated) ||
                    (!variation.HasValue && !string.IsNullOrWhiteSpace(goal.VariationAcceptableAsString));

                rows.Add(new ReviewPqmRow
                {
                    StableId = ClinicalReviewValueMapper.CreateUniqueStableId(
                        "eclipse-goal-" + structure + "-" + objective,
                        "eclipse-goal-" + (rows.Count + 1), usedIds),
                    TemplateCode = "ECLIPSE-" + (rows.Count + 1).ToString(CultureInfo.InvariantCulture),
                    TemplateStructure = structure,
                    ResolvedStructureId = matched ? structure : string.Empty,
                    // A native evaluation belongs to its assigned structure; do
                    // not offer alias reassignment of an already-evaluated goal.
                    StructureOptions = matched ? new List<string> { structure } : new List<string>(),
                    Objective = string.IsNullOrWhiteSpace(metric) ? goal.MeasureType.ToString() : metric,
                    Comparator = comparator,
                    Goal = limit,
                    Variation = variation,
                    AchievedValue = actual,
                    Unit = unit,
                    Status = status.Status,
                    Severity = status.Severity,
                    SourceLabel = "Eclipse Clinical Goals: assigned plan objectives",
                    MappingDescription = matched ? "Native assignment" : "Unresolved or empty",
                    Explanation = Safe(string.Format(CultureInfo.InvariantCulture,
                        "Eclipse Clinical Goals; structure: {0} ({1}); objective: {2}; " +
                        "achieved: {3}; variation: {4}; native evaluation: {5}.{6}",
                        structure, matched ? "exact match" : "unresolved or empty", objective,
                        Safe(goal.ActualValueAsString, "not available"),
                        Safe(goal.VariationAcceptableAsString, "not specified"),
                        goal.EvaluationResult, incomplete
                            ? " Numeric projection incomplete: explicit compatible units required; native text retained."
                            : string.Empty), "Native Eclipse clinical goal.")
                });
            }

            source = rows.Count == 0
                ? CreateSource(ReviewStatusCodes.NotConfigured,
                    "No Eclipse clinical goals are assigned to the active planning item; no successful evaluation is inferred.")
                : CreateSource(ReviewStatusCodes.Available,
                    rows.Count.ToString(CultureInfo.InvariantCulture) +
                    " native Eclipse clinical goal(s) read, separately from RefDB constraints. " +
                    rows.Count(item => item.Status == ReviewStatusCodes.NotEvaluated)
                        .ToString(CultureInfo.InvariantCulture) + " row(s) not evaluated or not structurally resolved.");
            return rows;
        }

        private static ReviewStatusAndSeverity MapStatus(GoalEvalResult status)
        {
            switch (status)
            {
                case GoalEvalResult.Passed:
                    return new ReviewStatusAndSeverity(ReviewStatusCodes.Pass, ReviewSeverityCodes.None);
                case GoalEvalResult.WithinVariationAcceptable:
                    return new ReviewStatusAndSeverity(ReviewStatusCodes.Variation, ReviewSeverityCodes.Warning);
                case GoalEvalResult.Failed:
                    return new ReviewStatusAndSeverity(ReviewStatusCodes.Fail, ReviewSeverityCodes.Error);
                default:
                    return new ReviewStatusAndSeverity(ReviewStatusCodes.NotEvaluated, ReviewSeverityCodes.Info);
            }
        }

        private static string MapComparator(ObjectiveOperator value)
        {
            switch (value)
            {
                case ObjectiveOperator.LessThan: return "<";
                case ObjectiveOperator.LessThanOrEqual: return "<=";
                case ObjectiveOperator.GreaterThan: return ">";
                case ObjectiveOperator.GreaterThanOrEqual: return ">=";
                case ObjectiveOperator.Equals: return "=";
                default: return string.Empty;
            }
        }

        private static double? ParseScalar(string text, out string unit)
        {
            unit = ReviewUnitCodes.Text;
            Match match = ScalarWithUnit.Match(text ?? string.Empty);
            double value;
            if (!match.Success || !double.TryParse(match.Groups["value"].Value.Replace(',', '.'),
                NumberStyles.Float, CultureInfo.InvariantCulture, out value) || !Finite(value).HasValue)
            {
                return null;
            }
            string explicitUnit = match.Groups["unit"].Value;
            if (string.Equals(explicitUnit, "cGy", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(explicitUnit, "Gy", StringComparison.OrdinalIgnoreCase))
            {
                unit = ReviewUnitCodes.Gray;
                return ClinicalReviewValueMapper.ConvertDoseToGray(value, explicitUnit);
            }
            unit = explicitUnit == "%" ? ReviewUnitCodes.Percent : ReviewUnitCodes.CubicCentimeter;
            return value;
        }

        private static double? Finite(double value)
        {
            return double.IsNaN(value) || double.IsInfinity(value) ? (double?)null : value;
        }

        private static string Safe(string value, string fallback)
        {
            return ClinicalReviewValueMapper.SanitizeClinicalLabel(value, fallback);
        }

        private static ReviewSourceStatus CreateSource(string status, string message)
        {
            return new ReviewSourceStatus
            {
                StableId = SourceStableId,
                SourceCode = "eclipse-clinical-goals",
                SourceType = "native-planning-item-goals",
                Status = status,
                Optional = true,
                UsedFallback = false,
                PathDisplayLabel = "Active Eclipse Clinical Goals",
                Message = message
            };
        }
    }
}
