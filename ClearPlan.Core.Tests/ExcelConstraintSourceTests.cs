using System;
using System.IO;
using System.Linq;
using ClearPlan.Core.Constraints;

namespace ClearPlan.Core.Tests
{
    internal static class ExcelConstraintSourceTests
    {
        public static void LoadsSharedStringsAndUnicode()
        {
            ConstraintCatalog catalog = new ExcelConstraintSource().Load(
                Fixture("excel-catalog-fixture.xlsx"));

            TestAssert.Equal("Excel", catalog.SourceKind);
            TestAssert.Equal(1, catalog.Tables.Count);
            TestAssert.Equal(2, catalog.Tables[0].Constraints.Count);
            TestAssert.Equal(2, catalog.Structures.Count);
            TestAssert.True(catalog.Structures.Single(
                item => item.StructureId == "SpinalCord").Aliases.Contains("Rückenmark"));
            TestAssert.True(catalog.Tables[0].Constraints.Any(
                item => item.Comment.Contains("Größe") &&
                        item.RawValues["legacy_metadata"].Contains("≤") &&
                        item.RawValues["legacy_metadata"].Contains("cm³")));
            TestAssert.False(catalog.Issues.Any(issue => issue.IsFatal));
        }

        public static void ReportsMissingHeaders()
        {
            ConstraintCatalog catalog = new ExcelConstraintSource().Load(
                Fixture("excel-catalog-missing-header.xlsx"));

            TestAssert.Equal(0, catalog.Tables.Count);
            TestAssert.True(catalog.Issues.Any(issue => issue.IsFatal &&
                                                        issue.Field == "missing_header" &&
                                                        issue.Message.Contains("table_id")));
        }

        public static void ReportsDuplicateIdsAndBrokenRelationships()
        {
            ConstraintCatalog catalog = new ExcelConstraintSource().Load(
                Fixture("excel-catalog-invalid-links.xlsx"));

            TestAssert.True(catalog.Issues.Any(issue => issue.IsFatal &&
                                                        issue.Field == "table_id" &&
                                                        issue.Message.Contains("duplicate")));
            TestAssert.True(catalog.Issues.Any(issue => issue.IsFatal &&
                                                        issue.Field == "constraint.table_id" &&
                                                        issue.Message.Contains("unknown")));
        }

        private static string Fixture(string fileName)
        {
            return Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Fixtures",
                fileName);
        }
    }
}
