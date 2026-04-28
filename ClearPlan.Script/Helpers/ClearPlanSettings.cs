using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace ClearPlan.Helpers
{
    public sealed class ClearPlanSettings
    {
        private const string SettingsJsonFileName = "settings.json";
        private const string LegacySettingsIniFileName = "settings.ini";
        private static readonly Lazy<ClearPlanSettings> CurrentSettings =
            new Lazy<ClearPlanSettings>(LoadInternal);

        public SettingsPathOptions Paths { get; set; } = new SettingsPathOptions();
        public SettingsLinkOptions Links { get; set; } = new SettingsLinkOptions();
        public SettingsCheckOptions Checks { get; set; } = new SettingsCheckOptions();

        public static ClearPlanSettings Load()
        {
            return CurrentSettings.Value;
        }

        public string BaseDirectory
        {
            get { return AssemblyHelper.GetAssemblyDirectory(); }
        }

        public string ResolvePath(string relativeOrAbsolutePath)
        {
            if (string.IsNullOrWhiteSpace(relativeOrAbsolutePath))
            {
                return BaseDirectory;
            }

            if (Path.IsPathRooted(relativeOrAbsolutePath))
            {
                return relativeOrAbsolutePath;
            }

            return Path.GetFullPath(Path.Combine(BaseDirectory, relativeOrAbsolutePath));
        }

        public string GetConstraintTemplatesDirectory()
        {
            return EnsureDirectory(ResolvePath(Paths.ConstraintTemplatesDirectory));
        }

        public string GetLogsDirectory()
        {
            return EnsureDirectory(ResolvePath(Paths.LogsDirectory));
        }

        public string GetReportsDirectory()
        {
            return EnsureDirectory(ResolvePath(Paths.ReportsDirectory));
        }

        public string GetExportsDirectory()
        {
            return EnsureDirectory(ResolvePath(Paths.CsvExportDirectory));
        }

        public string GetStateDirectory()
        {
            return EnsureDirectory(ResolvePath(Paths.StateDirectory));
        }

        public string GetUsageLogFile()
        {
            return EnsureParentDirectory(ResolvePath(Paths.UsageLogFile));
        }

        public string GetActivityLogFile()
        {
            return EnsureParentDirectory(ResolvePath(Paths.ActivityLogFile));
        }

        public string GetVersionSeenUsersFile()
        {
            return EnsureParentDirectory(ResolvePath(Paths.VersionSeenUsersFile));
        }

        public string GetConstraintTemplatePath(string fileName)
        {
            return Path.Combine(GetConstraintTemplatesDirectory(), fileName);
        }

        public string GetChangeLogFile()
        {
            return ResolvePath(Paths.ChangeLogFile);
        }

        public string GetFeedbackFile()
        {
            return ResolvePath(Paths.FeedbackFile);
        }

        private static ClearPlanSettings LoadInternal()
        {
            var settings = new ClearPlanSettings();
            settings.ApplyDefaults();

            string settingsJsonPath = Path.Combine(AssemblyHelper.GetAssemblyDirectory(), SettingsJsonFileName);
            if (File.Exists(settingsJsonPath))
            {
                try
                {
                    settings = JsonConvert.DeserializeObject<ClearPlanSettings>(File.ReadAllText(settingsJsonPath))
                        ?? new ClearPlanSettings();
                }
                catch
                {
                    settings = new ClearPlanSettings();
                }
            }
            else
            {
                ApplyLegacyIniSettings(settings);
            }

            settings.ApplyDefaults();
            return settings;
        }

        private static void ApplyLegacyIniSettings(ClearPlanSettings settings)
        {
            string iniPath = Path.Combine(AssemblyHelper.GetAssemblyDirectory(), LegacySettingsIniFileName);
            if (!File.Exists(iniPath))
            {
                return;
            }

            var legacyValues = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string activeSection = string.Empty;

            foreach (string rawLine in File.ReadAllLines(iniPath))
            {
                string line = rawLine.Trim();
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith(";"))
                {
                    continue;
                }

                if (line.StartsWith("[") && line.EndsWith("]"))
                {
                    activeSection = line.Substring(1, line.Length - 2);
                    continue;
                }

                if (!string.Equals(activeSection, "Paths", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string[] parts = line.Split(new[] { '=' }, 2);
                if (parts.Length == 2)
                {
                    legacyValues[parts[0].Trim()] = parts[1].Trim();
                }
            }

            string value;
            if (legacyValues.TryGetValue("ConstraintTemplatesDir", out value) && !string.IsNullOrWhiteSpace(value))
            {
                settings.Paths.ConstraintTemplatesDirectory = value;
            }

            if (legacyValues.TryGetValue("TConvFileName", out value) && !string.IsNullOrWhiteSpace(value))
            {
                settings.Paths.DefaultConventionalTemplate = value;
            }

            if (legacyValues.TryGetValue("UserLogPath", out value) && !string.IsNullOrWhiteSpace(value))
            {
                settings.Paths.UsageLogFile = value;
            }

            if (legacyValues.TryGetValue("UserListPath", out value) && !string.IsNullOrWhiteSpace(value))
            {
                settings.Paths.VersionSeenUsersFile = value;
            }

            if (legacyValues.TryGetValue("ReportsPath", out value) && !string.IsNullOrWhiteSpace(value))
            {
                settings.Paths.ReportsDirectory = value;
            }
        }

        private void ApplyDefaults()
        {
            if (Paths == null)
            {
                Paths = new SettingsPathOptions();
            }

            if (Links == null)
            {
                Links = new SettingsLinkOptions();
            }

            if (Checks == null)
            {
                Checks = new SettingsCheckOptions();
            }

            Paths.ConstraintTemplatesDirectory = DefaultIfBlank(Paths.ConstraintTemplatesDirectory, "ConstraintTemplates");
            Paths.DefaultConventionalTemplate = DefaultIfBlank(Paths.DefaultConventionalTemplate, "Starter_Conventional.csv");
            Paths.DefaultHypofractionatedTemplate = DefaultIfBlank(Paths.DefaultHypofractionatedTemplate, "Starter_Hypofractionated.csv");
            Paths.DefaultPlanSumTemplate = DefaultIfBlank(Paths.DefaultPlanSumTemplate, "Starter_PlanSum.csv");
            Paths.LogsDirectory = DefaultIfBlank(Paths.LogsDirectory, "Logs");
            Paths.ReportsDirectory = DefaultIfBlank(Paths.ReportsDirectory, "Reports");
            Paths.CsvExportDirectory = DefaultIfBlank(Paths.CsvExportDirectory, "Exports");
            Paths.StateDirectory = DefaultIfBlank(Paths.StateDirectory, "State");
            Paths.UsageLogFile = DefaultIfBlank(Paths.UsageLogFile, Path.Combine("Logs", "ClearPlan_UserLog.csv"));
            Paths.ActivityLogFile = DefaultIfBlank(Paths.ActivityLogFile, Path.Combine("Logs", "ActivityLog.csv"));
            Paths.VersionSeenUsersFile = DefaultIfBlank(Paths.VersionSeenUsersFile, Path.Combine("State", "seen_versions.csv"));
            Paths.ChangeLogFile = DefaultIfBlank(Paths.ChangeLogFile, "CHANGELOG.md");
            Paths.FeedbackFile = DefaultIfBlank(Paths.FeedbackFile, "FEEDBACK.md");

            Links.FeedbackUrl = Links.FeedbackUrl ?? string.Empty;
            Checks.Profile = DefaultIfBlank(Checks.Profile, "starter");
        }

        private static string DefaultIfBlank(string value, string fallbackValue)
        {
            return string.IsNullOrWhiteSpace(value) ? fallbackValue : value;
        }

        private static string EnsureDirectory(string directoryPath)
        {
            Directory.CreateDirectory(directoryPath);
            return directoryPath;
        }

        private static string EnsureParentDirectory(string filePath)
        {
            string directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            return filePath;
        }
    }

    public sealed class SettingsPathOptions
    {
        public string ConstraintTemplatesDirectory { get; set; }
        public string DefaultConventionalTemplate { get; set; }
        public string DefaultHypofractionatedTemplate { get; set; }
        public string DefaultPlanSumTemplate { get; set; }
        public string LogsDirectory { get; set; }
        public string ReportsDirectory { get; set; }
        public string CsvExportDirectory { get; set; }
        public string StateDirectory { get; set; }
        public string UsageLogFile { get; set; }
        public string ActivityLogFile { get; set; }
        public string VersionSeenUsersFile { get; set; }
        public string ChangeLogFile { get; set; }
        public string FeedbackFile { get; set; }
    }

    public sealed class SettingsLinkOptions
    {
        public string FeedbackUrl { get; set; }
    }

    public sealed class SettingsCheckOptions
    {
        public string Profile { get; set; }
    }
}
