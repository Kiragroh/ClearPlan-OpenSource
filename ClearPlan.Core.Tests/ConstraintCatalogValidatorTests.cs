using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class ConstraintCatalogValidatorTests
    {
        public static void RejectsComparatorAgainstMetricDirection()
        {
            var lowerIsBetter = new ConstraintDefinition
            {
                ConstraintId = "c1",
                TableId = "t1",
                StructureId = "SpinalCord",
                Metric = "D0.03cc",
                Unit = "Gy",
                Comparator = ">=",
                ExpectedDirection = ConstraintDirection.LowerIsBetter,
                Goal = 40m
            };
            var higherIsBetter = new ConstraintDefinition
            {
                ConstraintId = "c2",
                TableId = "t1",
                StructureId = "Liver",
                Metric = "CV21.5Gy",
                Unit = "cc",
                Comparator = "<=",
                ExpectedDirection = ConstraintDirection.HigherIsBetter,
                Goal = 700m
            };

            var issues = ConstraintCatalogValidator.ValidateConstraints(
                new[] { lowerIsBetter, higherIsBetter });

            TestAssert.Equal(2, issues.Count(issue => issue.Severity == CatalogIssueSeverity.Error));
            TestAssert.True(issues.All(issue => issue.Field == "comparator"));
        }

        public static void ReportsDuplicateAndMissingIdentities()
        {
            var catalog = new ConstraintCatalog
            {
                SourceKind = "Test",
                SourcePath = "fixture",
                Tables =
                {
                    new ConstraintTableDefinition { TableId = "table-a", DisplayName = "A", Active = true },
                    new ConstraintTableDefinition { TableId = "table-a", DisplayName = "Duplicate", Active = true }
                },
                Structures =
                {
                    new StructureDefinition { StructureId = "", CanonicalName = "Missing ID", Active = true }
                }
            };

            var issues = ConstraintCatalogValidator.Validate(catalog);

            TestAssert.True(issues.Any(issue => issue.Field == "table_id" &&
                                                issue.Message.Contains("duplicate")));
            TestAssert.True(issues.Any(issue => issue.Field == "structure_id" &&
                                                issue.Message.Contains("required")));
        }
    }
}
