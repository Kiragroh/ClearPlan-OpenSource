using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClearPlan.Core.Constraints
{
    public static class PqmObjectiveMapper
    {
        public static PqmObjectiveDefinition Map(ConstraintDefinition constraint)
        {
            if (constraint == null)
            {
                throw new ArgumentNullException("constraint");
            }

            string templateId = FirstNonEmpty(
                constraint.StructureName,
                constraint.RawStructureName,
                constraint.StructureId,
                constraint.ConstraintId);
            var objective = new PqmObjectiveDefinition
            {
                ConstraintId = constraint.ConstraintId,
                TemplateId = templateId,
                DvhObjective = FirstNonEmpty(
                    constraint.DvhObjective,
                    BuildDvhObjective(constraint.Metric, constraint.Unit)),
                Goal = FirstNonEmpty(
                    constraint.EvaluationPoint,
                    BuildEvaluationPoint(constraint.Comparator, constraint.Goal)),
                Variation = FormatDecimal(constraint.Variation),
                Priority = constraint.Priority.HasValue
                    ? constraint.Priority.Value.ToString(CultureInfo.InvariantCulture)
                    : string.Empty,
                Source = constraint.Source ?? string.Empty,
                Comment = constraint.Comment ?? string.Empty
            };
            objective.Aliases = CopyOrFallback(constraint.Aliases, templateId);
            objective.Codes = CopyOrFallback(constraint.Codes, templateId);
            objective.DicomTypes = CopyOrFallback(constraint.DicomTypes, templateId);
            return objective;
        }

        private static IList<string> CopyOrFallback(
            IEnumerable<string> values,
            string fallback)
        {
            IList<string> result = (values ?? Enumerable.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value))
                .Select(value => value.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (result.Count == 0 && !string.IsNullOrWhiteSpace(fallback))
            {
                result.Add(fallback);
            }

            return result;
        }

        private static string BuildDvhObjective(string metric, string unit)
        {
            string objective = metric ?? string.Empty;
            if (objective.Equals("Dmean", StringComparison.OrdinalIgnoreCase))
            {
                objective = "Mean";
            }
            else if (objective.Equals("Dmax", StringComparison.OrdinalIgnoreCase))
            {
                objective = "Max";
            }
            else if (objective.Equals("Dmin", StringComparison.OrdinalIgnoreCase))
            {
                objective = "Min";
            }

            return string.IsNullOrWhiteSpace(unit)
                ? objective
                : string.Format("{0}[{1}]", objective, unit);
        }

        private static string BuildEvaluationPoint(
            string comparator,
            decimal? goal)
        {
            return (comparator ?? string.Empty) + FormatDecimal(goal);
        }

        private static string FormatDecimal(decimal? value)
        {
            return value.HasValue
                ? value.Value.ToString("0.###", CultureInfo.InvariantCulture)
                : string.Empty;
        }

        private static string FirstNonEmpty(params string[] values)
        {
            return values.FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
                   ?? string.Empty;
        }
    }
}
