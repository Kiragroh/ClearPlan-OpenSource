using ClearPlan.Core.Settings;

namespace ClearPlan.Core.Tests
{
    internal static class PathSettingsIniTests
    {
        public static void RoundTripsEveryConfiguredPath()
        {
            var ratePath = typeof(ClearPlanPathOptions).GetProperty("DoseRateProfilesJsonPath");
            TestAssert.NotNull(ratePath, "Dose-rate assumptions require an editable external profile path.");
            ClearPlanSettingsModel source = SettingsWithUniquePaths();
            ratePath.SetValue(source.Paths, "MachineGeometry/rate-estimates.json");

            string ini = PathSettingsIni.Serialize(source);
            var restored = new ClearPlanSettingsModel();
            PathSettingsIni.Apply(restored, ini);
            TestAssert.Equal("MachineGeometry/rate-estimates.json", (string)ratePath.GetValue(restored.Paths));

            TestAssert.Equal(
                source.ConstraintSource.RefDbJsonPath,
                restored.ConstraintSource.RefDbJsonPath);
            TestAssert.Equal(
                source.ConstraintSource.ExcelWorkbookPath,
                restored.ConstraintSource.ExcelWorkbookPath);
            TestAssert.Equal(
                source.Paths.ConstraintTemplatesDirectory,
                restored.Paths.ConstraintTemplatesDirectory);
            TestAssert.Equal(
                source.Paths.DefaultConventionalTemplate,
                restored.Paths.DefaultConventionalTemplate);
            TestAssert.Equal(
                source.Paths.DefaultHypofractionatedTemplate,
                restored.Paths.DefaultHypofractionatedTemplate);
            TestAssert.Equal(
                source.Paths.DefaultPlanSumTemplate,
                restored.Paths.DefaultPlanSumTemplate);
            TestAssert.Equal(
                source.Paths.LogsDirectory,
                restored.Paths.LogsDirectory);
            TestAssert.Equal(
                source.Paths.ReportsDirectory,
                restored.Paths.ReportsDirectory);
            TestAssert.Equal(
                source.Paths.CsvExportDirectory,
                restored.Paths.CsvExportDirectory);
            TestAssert.Equal(
                source.Paths.StateDirectory,
                restored.Paths.StateDirectory);
            TestAssert.Equal(
                source.Paths.UsageLogFile,
                restored.Paths.UsageLogFile);
            TestAssert.Equal(
                source.Paths.ActivityLogFile,
                restored.Paths.ActivityLogFile);
            TestAssert.Equal(
                source.Paths.VersionSeenUsersFile,
                restored.Paths.VersionSeenUsersFile);
            TestAssert.Equal(
                source.Paths.ChangeLogFile,
                restored.Paths.ChangeLogFile);
            TestAssert.Equal(
                source.Paths.FeedbackFile,
                restored.Paths.FeedbackFile);
            TestAssert.Equal(
                source.ClinicalPaths.PrescriptionSettingsPath,
                restored.ClinicalPaths.PrescriptionSettingsPath);
            TestAssert.Equal(
                source.ClinicalPaths.PlanCheckResourcesPath,
                restored.ClinicalPaths.PlanCheckResourcesPath);
        }

        public static void ReadsLegacyPathAliases()
        {
            var aria=new ClearPlanSettingsModel();
            aria.Paths.AriaUploadConfigJsonPath="Local/aria-connection.json";
            var ariaRestored=new ClearPlanSettingsModel(); PathSettingsIni.Apply(ariaRestored,PathSettingsIni.Serialize(aria));
            TestAssert.Equal(aria.Paths.AriaUploadConfigJsonPath,ariaRestored.Paths.AriaUploadConfigJsonPath);
            var settings = new ClearPlanSettingsModel();
            const string ini =
                "[Paths]\n" +
                "ConstraintTemplatesDir = LegacyTemplates\n" +
                "UserLogPath = LegacyLogs/User.csv\n" +
                "UserListPath = LegacyState/seen.csv\n" +
                "ReportsPath = LegacyReports\n" +
                "TConvFileName = LegacyConv.xlsx\n";

            PathSettingsIni.Apply(settings, ini);

            TestAssert.Equal(
                "LegacyTemplates",
                settings.Paths.ConstraintTemplatesDirectory);
            TestAssert.Equal(
                "LegacyLogs/User.csv",
                settings.Paths.UsageLogFile);
            TestAssert.Equal(
                "LegacyState/seen.csv",
                settings.Paths.VersionSeenUsersFile);
            TestAssert.Equal(
                "LegacyReports",
                settings.Paths.ReportsDirectory);
            TestAssert.Equal(
                "LegacyConv.xlsx",
                settings.Paths.DefaultConventionalTemplate);
        }

        private static ClearPlanSettingsModel SettingsWithUniquePaths()
        {
            var settings = new ClearPlanSettingsModel();
            settings.ConstraintSource.RefDbJsonPath = "RefDB/strukturen-ä.json";
            settings.ConstraintSource.ExcelWorkbookPath = "Catalog/catalog.xlsx";
            settings.Paths.ConstraintTemplatesDirectory = "Templates";
            settings.Paths.DefaultConventionalTemplate = "conv.xlsx";
            settings.Paths.DefaultHypofractionatedTemplate = "hypo.xlsx";
            settings.Paths.DefaultPlanSumTemplate = "sum.xlsx";
            settings.Paths.LogsDirectory = "Logs";
            settings.Paths.ReportsDirectory = "Reports";
            settings.Paths.CsvExportDirectory = "Exports";
            settings.Paths.StateDirectory = "State";
            settings.Paths.UsageLogFile = "Logs/user.csv";
            settings.Paths.ActivityLogFile = "Logs/activity.csv";
            settings.Paths.VersionSeenUsersFile = "State/seen.csv";
            settings.Paths.ChangeLogFile = "CHANGELOG.md";
            settings.Paths.FeedbackFile = "FEEDBACK.md";
            settings.ClinicalPaths.PrescriptionSettingsPath = "Clinical/Rx";
            settings.ClinicalPaths.PlanCheckResourcesPath = "Clinical/PlanCheck";
            return settings;
        }
    }
}
