using System;
using System.IO;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class SettingsValidatorTests
    {
        public static void RequiresConfiguredSourceByMode()
        {
            string baseDirectory = CreateTemporaryDirectory();
            try
            {
                ClearPlanSettingsModel settings = ValidSettings(baseDirectory);
                settings.ConstraintSource.Mode = ConstraintSourceMode.RefDb;
                settings.ConstraintSource.RefDbJsonPath = string.Empty;

                SettingsValidationResult result =
                    SettingsValidator.Validate(settings, baseDirectory);

                TestAssert.False(result.IsValid);
                TestAssert.True(result.Errors.Exists(
                    value => value.IndexOf("RefDB", StringComparison.OrdinalIgnoreCase) >= 0));
            }
            finally
            {
                Directory.Delete(baseDirectory, true);
            }
        }

        public static void ChecksSourceExtensionsAndReadability()
        {
            string baseDirectory = CreateTemporaryDirectory();
            try
            {
                ClearPlanSettingsModel settings = ValidSettings(baseDirectory);
                string wrongExtension = Path.Combine(baseDirectory, "catalog.txt");
                File.WriteAllText(wrongExtension, "{}");
                settings.ConstraintSource.Mode = ConstraintSourceMode.RefDb;
                settings.ConstraintSource.RefDbJsonPath = wrongExtension;

                SettingsValidationResult result =
                    SettingsValidator.Validate(settings, baseDirectory);

                TestAssert.False(result.IsValid);
                TestAssert.True(result.Errors.Exists(
                    value => value.IndexOf(".json", StringComparison.OrdinalIgnoreCase) >= 0));
            }
            finally
            {
                Directory.Delete(baseDirectory, true);
            }
        }

        public static void AcceptsWritableOperationalDirectories()
        {
            string baseDirectory = CreateTemporaryDirectory();
            try
            {
                ClearPlanSettingsModel settings = ValidSettings(baseDirectory);

                SettingsValidationResult result =
                    SettingsValidator.Validate(settings, baseDirectory);

                TestAssert.True(result.IsValid, string.Join("; ", result.Errors));
                TestAssert.Equal(0, result.Errors.Count);
            }
            finally
            {
                Directory.Delete(baseDirectory, true);
            }
        }

        private static ClearPlanSettingsModel ValidSettings(string baseDirectory)
        {
            string workbook = Path.Combine(baseDirectory, "catalog.xlsx");
            File.WriteAllBytes(workbook, new byte[] { 1, 2, 3 });
            return new ClearPlanSettingsModel
            {
                ConstraintSource = new ConstraintSourceOptions
                {
                    Mode = ConstraintSourceMode.Excel,
                    ExcelWorkbookPath = workbook
                },
                Paths = new ClearPlanPathOptions
                {
                    LogsDirectory = Path.Combine(baseDirectory, "Logs"),
                    ReportsDirectory = Path.Combine(baseDirectory, "Reports"),
                    CsvExportDirectory = Path.Combine(baseDirectory, "Exports"),
                    StateDirectory = Path.Combine(baseDirectory, "State")
                }
            };
        }

        private static string CreateTemporaryDirectory()
        {
            string path = Path.Combine(
                Path.GetTempPath(),
                "ClearPlanSettingsTests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(path);
            return path;
        }
    }
}
