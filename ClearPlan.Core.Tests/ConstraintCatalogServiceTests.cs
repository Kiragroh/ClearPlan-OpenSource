using System;
using System.IO;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class ConstraintCatalogServiceTests
    {
        public static void AutomaticUsesValidRefDbFirst()
        {
            var options = new ConstraintSourceOptions
            {
                Mode = ConstraintSourceMode.Automatic,
                RefDbJsonPath = "refdb-minimal.json",
                ExcelWorkbookPath = "excel-catalog-fixture.xlsx"
            };

            ConstraintCatalogLoadResult result = new ConstraintCatalogService().Load(
                options,
                FixtureDirectory());

            TestAssert.True(result.IsUsable);
            TestAssert.Equal("RefDB", result.ActiveSource);
            TestAssert.Equal(1, result.Catalog.Tables.Count);
            TestAssert.Equal(0, result.Errors.Count);
        }

        public static void AutomaticFallsBackToExcelWithWarning()
        {
            var options = new ConstraintSourceOptions
            {
                Mode = ConstraintSourceMode.Automatic,
                RefDbJsonPath = "missing-refdb.json",
                ExcelWorkbookPath = "excel-catalog-fixture.xlsx"
            };

            ConstraintCatalogLoadResult result = new ConstraintCatalogService().Load(
                options,
                FixtureDirectory());

            TestAssert.True(result.IsUsable);
            TestAssert.Equal("Excel", result.ActiveSource);
            TestAssert.True(result.Warnings.Count > 0);
            TestAssert.Equal(0, result.Errors.Count);
        }

        public static void ExplicitModesNeverSilentlyFallback()
        {
            var refDbOnly = new ConstraintSourceOptions
            {
                Mode = ConstraintSourceMode.RefDb,
                RefDbJsonPath = "missing-refdb.json",
                ExcelWorkbookPath = "excel-catalog-fixture.xlsx"
            };
            ConstraintCatalogLoadResult refDbResult = new ConstraintCatalogService().Load(
                refDbOnly,
                FixtureDirectory());
            TestAssert.False(refDbResult.IsUsable);
            TestAssert.Equal(string.Empty, refDbResult.ActiveSource);
            TestAssert.True(refDbResult.Errors.Count > 0);

            var excelOnly = new ConstraintSourceOptions
            {
                Mode = ConstraintSourceMode.Excel,
                RefDbJsonPath = "refdb-minimal.json",
                ExcelWorkbookPath = "excel-catalog-fixture.xlsx"
            };
            ConstraintCatalogLoadResult excelResult = new ConstraintCatalogService().Load(
                excelOnly,
                FixtureDirectory());
            TestAssert.True(excelResult.IsUsable);
            TestAssert.Equal("Excel", excelResult.ActiveSource);
        }

        public static void AutomaticBlocksWhenNoSourceIsUsable()
        {
            var options = new ConstraintSourceOptions
            {
                Mode = ConstraintSourceMode.Automatic,
                RefDbJsonPath = "missing-refdb.json",
                ExcelWorkbookPath = "missing-workbook.xlsx"
            };

            ConstraintCatalogLoadResult result = new ConstraintCatalogService().Load(
                options,
                FixtureDirectory());

            TestAssert.False(result.IsUsable);
            TestAssert.True(result.Errors.Count > 0);
            TestAssert.Equal(string.Empty, result.ActiveSource);
        }

        public static void ResolvesRelativeAndAbsolutePaths()
        {
            string baseDirectory = FixtureDirectory();
            string relative = SettingsPathResolver.Resolve(
                baseDirectory,
                "excel-catalog-fixture.xlsx");
            TestAssert.Equal(
                Path.Combine(baseDirectory, "excel-catalog-fixture.xlsx"),
                relative);

            string absolute = Path.Combine(baseDirectory, "refdb-minimal.json");
            TestAssert.Equal(
                absolute,
                SettingsPathResolver.Resolve(baseDirectory, absolute));
        }

        private static string FixtureDirectory()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Fixtures");
        }
    }
}
