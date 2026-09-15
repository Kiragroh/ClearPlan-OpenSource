using System;
using System.Collections.Generic;
using System.Text;

namespace ClearPlan.Core.Settings
{
    public static class PathSettingsIni
    {
        public static void Apply(
            ClearPlanSettingsModel settings,
            string iniText)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            EnsureSections(settings);
            IDictionary<string, string> values = ParsePaths(iniText);
            Apply(values, "MlcGeometryProfilesJsonPath",
                value => settings.Paths.MlcGeometryProfilesJsonPath = value);
            Apply(values, "DoseRateProfilesJsonPath",
                value => settings.Paths.DoseRateProfilesJsonPath = value);
            Apply(values, "CollisionProfilesJsonPath",
                value => settings.Paths.CollisionProfilesJsonPath = value);
            Apply(values, "SourceCollisionModelsJsonPath",
                value => settings.Paths.SourceCollisionModelsJsonPath = value);
            Apply(values, "AriaUploadConfigJsonPath",
                value => settings.Paths.AriaUploadConfigJsonPath = value);
            Apply(values, "DefaultReviewRulesJsonPath",
                value => settings.Paths.DefaultReviewRulesJsonPath = value);
            Apply(values, "PlanCheckSelectionJsonPath",
                value => settings.Paths.PlanCheckSelectionJsonPath = value);
            Apply(values, "FieldNamingRulesJsonPath",
                value => settings.Paths.FieldNamingRulesJsonPath = value);
            Apply(values, "IsodoseDisplayJsonPath",
                value => settings.Paths.IsodoseDisplayJsonPath = value);
            Apply(values, "RefDbJsonPath",
                value => settings.ConstraintSource.RefDbJsonPath = value);
            Apply(values, "ExcelWorkbookPath",
                value => settings.ConstraintSource.ExcelWorkbookPath = value);
            Apply(values, "StructureAliasesJsonPath",
                value => settings.ConstraintSource.StructureAliasesJsonPath = value);
            Apply(values, "ConstraintTemplatesDirectory",
                value => settings.Paths.ConstraintTemplatesDirectory = value);
            Apply(values, "DefaultConventionalTemplate",
                value => settings.Paths.DefaultConventionalTemplate = value);
            Apply(values, "DefaultHypofractionatedTemplate",
                value => settings.Paths.DefaultHypofractionatedTemplate = value);
            Apply(values, "DefaultPlanSumTemplate",
                value => settings.Paths.DefaultPlanSumTemplate = value);
            Apply(values, "LogsDirectory",
                value => settings.Paths.LogsDirectory = value);
            Apply(values, "ReportsDirectory",
                value => settings.Paths.ReportsDirectory = value);
            Apply(values, "CsvExportDirectory",
                value => settings.Paths.CsvExportDirectory = value);
            Apply(values, "StateDirectory",
                value => settings.Paths.StateDirectory = value);
            Apply(values, "ConfigurationDirectory",
                value => settings.Paths.ConfigurationDirectory = value);
            Apply(values, "UsageLogFile",
                value => settings.Paths.UsageLogFile = value);
            Apply(values, "ActivityLogFile",
                value => settings.Paths.ActivityLogFile = value);
            Apply(values, "VersionSeenUsersFile",
                value => settings.Paths.VersionSeenUsersFile = value);
            Apply(values, "ChangeLogFile",
                value => settings.Paths.ChangeLogFile = value);
            Apply(values, "FeedbackFile",
                value => settings.Paths.FeedbackFile = value);
            Apply(values, "PrescriptionSettingsPath",
                value => settings.ClinicalPaths.PrescriptionSettingsPath = value);
            Apply(values, "PlanCheckResourcesPath",
                value => settings.ClinicalPaths.PlanCheckResourcesPath = value);

            ApplyLegacyAliases(settings, values);
        }

        public static string Serialize(ClearPlanSettingsModel settings)
        {
            if (settings == null)
            {
                throw new ArgumentNullException("settings");
            }

            EnsureSections(settings);
            var builder = new StringBuilder();
            builder.AppendLine("; ClearPlan path settings");
            builder.AppendLine("; Relative paths are resolved from the build folder.");
            builder.AppendLine();
            builder.AppendLine("[Paths]");
            Append(builder, "MlcGeometryProfilesJsonPath", settings.Paths.MlcGeometryProfilesJsonPath);
            Append(builder, "DoseRateProfilesJsonPath", settings.Paths.DoseRateProfilesJsonPath);
            Append(builder, "CollisionProfilesJsonPath", settings.Paths.CollisionProfilesJsonPath);
            Append(builder, "SourceCollisionModelsJsonPath", settings.Paths.SourceCollisionModelsJsonPath);
            Append(builder, "AriaUploadConfigJsonPath", settings.Paths.AriaUploadConfigJsonPath);
            Append(builder, "DefaultReviewRulesJsonPath", settings.Paths.DefaultReviewRulesJsonPath);
            Append(builder, "PlanCheckSelectionJsonPath", settings.Paths.PlanCheckSelectionJsonPath);
            Append(builder, "FieldNamingRulesJsonPath", settings.Paths.FieldNamingRulesJsonPath);
            Append(builder, "IsodoseDisplayJsonPath", settings.Paths.IsodoseDisplayJsonPath);
            Append(builder, "RefDbJsonPath",
                settings.ConstraintSource.RefDbJsonPath);
            Append(builder, "ExcelWorkbookPath",
                settings.ConstraintSource.ExcelWorkbookPath);
            Append(builder, "StructureAliasesJsonPath",
                settings.ConstraintSource.StructureAliasesJsonPath);
            Append(builder, "ConstraintTemplatesDirectory",
                settings.Paths.ConstraintTemplatesDirectory);
            Append(builder, "DefaultConventionalTemplate",
                settings.Paths.DefaultConventionalTemplate);
            Append(builder, "DefaultHypofractionatedTemplate",
                settings.Paths.DefaultHypofractionatedTemplate);
            Append(builder, "DefaultPlanSumTemplate",
                settings.Paths.DefaultPlanSumTemplate);
            Append(builder, "LogsDirectory", settings.Paths.LogsDirectory);
            Append(builder, "ReportsDirectory", settings.Paths.ReportsDirectory);
            Append(builder, "CsvExportDirectory",
                settings.Paths.CsvExportDirectory);
            Append(builder, "StateDirectory", settings.Paths.StateDirectory);
            Append(builder, "ConfigurationDirectory", settings.Paths.ConfigurationDirectory);
            Append(builder, "UsageLogFile", settings.Paths.UsageLogFile);
            Append(builder, "ActivityLogFile", settings.Paths.ActivityLogFile);
            Append(builder, "VersionSeenUsersFile",
                settings.Paths.VersionSeenUsersFile);
            Append(builder, "ChangeLogFile", settings.Paths.ChangeLogFile);
            Append(builder, "FeedbackFile", settings.Paths.FeedbackFile);
            Append(builder, "PrescriptionSettingsPath",
                settings.ClinicalPaths.PrescriptionSettingsPath);
            Append(builder, "PlanCheckResourcesPath",
                settings.ClinicalPaths.PlanCheckResourcesPath);
            return builder.ToString();
        }

        private static IDictionary<string, string> ParsePaths(string iniText)
        {
            var values = new Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            string activeSection = string.Empty;
            string[] lines = (iniText ?? string.Empty).Replace(
                "\r\n",
                "\n").Split('\n');
            foreach (string rawLine in lines)
            {
                string line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line) ||
                    line.StartsWith(";") ||
                    line.StartsWith("#"))
                {
                    continue;
                }

                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    activeSection = line.Substring(1, line.Length - 2);
                    continue;
                }

                if (!string.Equals(
                    activeSection,
                    "Paths",
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string[] parts = line.Split(new[] { '=' }, 2);
                if (parts.Length == 2)
                {
                    values[parts[0].Trim()] = parts[1].Trim();
                }
            }

            return values;
        }

        private static void ApplyLegacyAliases(
            ClearPlanSettingsModel settings,
            IDictionary<string, string> values)
        {
            ApplyWhenCanonicalMissing(
                values,
                "ConstraintTemplatesDirectory",
                "ConstraintTemplatesDir",
                value => settings.Paths.ConstraintTemplatesDirectory = value);
            ApplyWhenCanonicalMissing(
                values,
                "UsageLogFile",
                "UserLogPath",
                value => settings.Paths.UsageLogFile = value);
            ApplyWhenCanonicalMissing(
                values,
                "VersionSeenUsersFile",
                "UserListPath",
                value => settings.Paths.VersionSeenUsersFile = value);
            ApplyWhenCanonicalMissing(
                values,
                "ReportsDirectory",
                "ReportsPath",
                value => settings.Paths.ReportsDirectory = value);
            ApplyWhenCanonicalMissing(
                values,
                "DefaultConventionalTemplate",
                "TConvFileName",
                value => settings.Paths.DefaultConventionalTemplate = value);
        }

        private static void ApplyWhenCanonicalMissing(
            IDictionary<string, string> values,
            string canonicalKey,
            string legacyKey,
            Action<string> assign)
        {
            if (values.ContainsKey(canonicalKey))
            {
                return;
            }

            Apply(values, legacyKey, assign);
        }

        private static void Apply(
            IDictionary<string, string> values,
            string key,
            Action<string> assign)
        {
            string value;
            if (values.TryGetValue(key, out value))
            {
                assign(value);
            }
        }

        private static void Append(
            StringBuilder builder,
            string key,
            string value)
        {
            builder.Append(key);
            builder.Append(" = ");
            builder.AppendLine(value ?? string.Empty);
        }

        private static void EnsureSections(ClearPlanSettingsModel settings)
        {
            settings.ConstraintSource = settings.ConstraintSource ??
                                        new ConstraintSourceOptions();
            settings.Paths = settings.Paths ?? new ClearPlanPathOptions();
            settings.ClinicalPaths = settings.ClinicalPaths ??
                                     new ClearPlanClinicalPathOptions();
        }
    }
}
