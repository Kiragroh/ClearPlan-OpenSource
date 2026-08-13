using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.Constraints
{
    public static class ConstraintCatalogValidator
    {
        public static IList<CatalogValidationIssue> Validate(ConstraintCatalog catalog)
        {
            var issues = new List<CatalogValidationIssue>();
            if (catalog == null)
            {
                issues.Add(Error("catalog", null, "Catalog is required.", "catalog"));
                return issues;
            }

            foreach (IGrouping<string, ConstraintTableDefinition> duplicate in catalog.Tables
                .Where(table => !string.IsNullOrWhiteSpace(table.TableId))
                .GroupBy(table => table.TableId, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1))
            {
                issues.Add(Error(
                    "table_id",
                    duplicate.Key,
                    "Table ID is duplicate.",
                    "Tables"));
            }

            foreach (ConstraintTableDefinition table in catalog.Tables)
            {
                if (string.IsNullOrWhiteSpace(table.TableId))
                {
                    issues.Add(Error("table_id", table.TableId, "Table ID is required.", "Tables"));
                }
            }

            foreach (StructureDefinition structure in catalog.Structures)
            {
                if (string.IsNullOrWhiteSpace(structure.StructureId))
                {
                    issues.Add(Error(
                        "structure_id",
                        structure.StructureId,
                        "Structure ID is required.",
                        "Structures"));
                }
            }

            foreach (CatalogValidationIssue issue in ValidateConstraints(
                catalog.Tables.SelectMany(table => table.Constraints)))
            {
                issues.Add(issue);
            }

            return issues;
        }

        public static IList<CatalogValidationIssue> ValidateConstraints(
            IEnumerable<ConstraintDefinition> constraints)
        {
            var issues = new List<CatalogValidationIssue>();
            foreach (ConstraintDefinition constraint in constraints ?? Enumerable.Empty<ConstraintDefinition>())
            {
                string comparator = ConstraintValueNormalizer.NormalizeComparator(constraint.Comparator);
                string metric = ConstraintValueNormalizer.NormalizeMetric(constraint.Metric);
                string location = string.Format(
                    "{0}/{1}",
                    constraint.TableId ?? "?",
                    constraint.ConstraintId ?? "?");

                if (constraint.ExpectedDirection == ConstraintDirection.HigherIsBetter &&
                    (comparator == "<" || comparator == "<="))
                {
                    issues.Add(Error(
                        "comparator",
                        constraint.Comparator,
                        "Comparator conflicts with the higher-is-better CV metric direction.",
                        location));
                }
                else if (constraint.ExpectedDirection == ConstraintDirection.LowerIsBetter &&
                         (comparator == ">" || comparator == ">="))
                {
                    issues.Add(Error(
                        "comparator",
                        constraint.Comparator,
                        "Comparator conflicts with the lower-is-better metric direction.",
                        location));
                }
            }

            return issues;
        }

        private static CatalogValidationIssue Error(
            string field,
            string rawValue,
            string message,
            string sourceLocation)
        {
            return new CatalogValidationIssue
            {
                Severity = CatalogIssueSeverity.Error,
                Field = field,
                RawValue = rawValue,
                Message = message,
                SourceLocation = sourceLocation
            };
        }
    }
}
