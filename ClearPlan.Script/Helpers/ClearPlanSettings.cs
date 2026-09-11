using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Settings;
using Newtonsoft.Json;

namespace ClearPlan.Helpers
{
    public sealed class ClearPlanSettings : ClearPlanSettingsModel
    {
        private const string SettingsJsonFileName = "settings.json";
        private const string PathSettingsIniFileName = "settings.ini";
        private static readonly TimeSpan NetworkSourceTimeout =
            TimeSpan.FromSeconds(2);
        private static readonly object SettingsLock = new object();
        private static ClearPlanSettings _current;
        private readonly object _constraintCatalogLock = new object();
        private ConstraintCatalogLoadResult _constraintCatalogCache;

        public static ClearPlanSettings Load()
        {
            lock (SettingsLock)
            {
                return _current ?? (_current = LoadInternal());
            }
        }

        public static ClearPlanSettings Reload()
        {
            lock (SettingsLock)
            {
                _current = LoadInternal();
                return _current;
            }
        }

        public string BaseDirectory
        {
            get { return AssemblyHelper.GetAssemblyDirectory(); }
        }

        public string SettingsFilePath
        {
            get { return Path.Combine(BaseDirectory, SettingsJsonFileName); }
        }

        public string PathSettingsFilePath
        {
            get { return Path.Combine(BaseDirectory, PathSettingsIniFileName); }
        }

        public string ResolvePath(string relativeOrAbsolutePath)
        {
            return SettingsPathResolver.Resolve(BaseDirectory, relativeOrAbsolutePath);
        }

        public ConstraintCatalogLoadResult LoadConstraintCatalog()
        {
            lock (_constraintCatalogLock)
            {
                if (_constraintCatalogCache == null)
                {
                    _constraintCatalogCache =
                        LoadConstraintCatalogWithNetworkBound();
                    ApplyConfiguredAliases(_constraintCatalogCache);
                }

                return _constraintCatalogCache;
            }
        }

        private ConstraintCatalogLoadResult
            LoadConstraintCatalogWithNetworkBound()
        {
            ConstraintSourceOptions source =
                ConstraintSource ?? new ConstraintSourceOptions();
            if (source.Mode != ConstraintSourceMode.Automatic)
            {
                return LoadConfiguredSource(
                    source.Mode,
                    source.Mode == ConstraintSourceMode.RefDb
                        ? source.RefDbJsonPath
                        : source.ExcelWorkbookPath,
                    source.IncludeInactiveTables);
            }

            ConstraintCatalogLoadResult refDb = LoadConfiguredSource(
                ConstraintSourceMode.RefDb,
                source.RefDbJsonPath,
                source.IncludeInactiveTables);
            if (refDb.IsUsable)
            {
                return refDb;
            }

            ConstraintCatalogLoadResult excel = LoadConfiguredSource(
                ConstraintSourceMode.Excel,
                source.ExcelWorkbookPath,
                source.IncludeInactiveTables);
            if (excel.IsUsable)
            {
                foreach (string warning in refDb.Warnings)
                {
                    excel.Warnings.Add("RefDB: " + warning);
                }
                foreach (string error in refDb.Errors)
                {
                    excel.Warnings.Add("RefDB: " + error);
                }
                excel.Warnings.Add(
                    "RefDB nicht verfügbar oder nicht rechtzeitig erreichbar – Excel aktiv.");
                return excel;
            }

            var failure = new ConstraintCatalogLoadResult
            {
                Catalog = excel.Catalog ?? refDb.Catalog
            };
            foreach (string warning in refDb.Warnings)
            {
                failure.Warnings.Add("RefDB: " + warning);
            }
            foreach (string warning in excel.Warnings)
            {
                failure.Warnings.Add("Excel: " + warning);
            }
            foreach (string error in refDb.Errors)
            {
                failure.Errors.Add("RefDB: " + error);
            }
            foreach (string error in excel.Errors)
            {
                failure.Errors.Add("Excel: " + error);
            }
            if (failure.Errors.Count == 0)
            {
                failure.Errors.Add(
                    "Neither RefDB nor Excel provides a usable constraint catalog.");
            }

            return failure;
        }

        private void ApplyConfiguredAliases(ConstraintCatalogLoadResult result)
        {
            string configuredPath = ConstraintSource == null ? null : ConstraintSource.StructureAliasesJsonPath;
            if (string.IsNullOrWhiteSpace(configuredPath) || !result.IsUsable) return;
            if (!IsUncPath(ResolvePath(configuredPath)))
            {
                ConstraintCatalogService.ApplyAliases(result, configuredPath, BaseDirectory);
                return;
            }
            // Work on a detached copy: a timed-out network read must never mutate the live catalog later.
            string detachedJson = JsonConvert.SerializeObject(result);
            var task = Task.Factory.StartNew(() =>
            {
                var copy = JsonConvert.DeserializeObject<ConstraintCatalogLoadResult>(detachedJson);
                ConstraintCatalogService.ApplyAliases(copy, configuredPath, BaseDirectory);
                return copy;
            }, CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);
            try
            {
                if (task.Wait(NetworkSourceTimeout))
                {
                    result.Catalog = task.Result.Catalog;
                    result.Warnings = task.Result.Warnings;
                    result.Errors = task.Result.Errors;
                    return;
                }
            }
            catch (AggregateException exception)
            {
                result.Errors.Add("Struktur-Aliase nicht angewendet: " + exception.GetBaseException().Message);
                return;
            }
            result.Errors.Add("Struktur-Aliase nicht angewendet: Netzwerkdatei nicht rechtzeitig erreichbar.");
            task.ContinueWith(completed => { var ignored = completed.Exception; }, TaskContinuationOptions.OnlyOnFaulted);
        }

        private ConstraintCatalogLoadResult LoadConfiguredSource(
            ConstraintSourceMode mode,
            string configuredPath,
            bool includeInactiveTables)
        {
            var options = new ConstraintSourceOptions
            {
                Mode = mode,
                RefDbJsonPath = configuredPath,
                ExcelWorkbookPath = configuredPath,
                IncludeInactiveTables = includeInactiveTables
            };
            var service = new ConstraintCatalogService();
            string resolvedConfiguredPath =
                ResolvePath(configuredPath);
            if (!IsUncPath(resolvedConfiguredPath))
            {
                return service.Load(options, BaseDirectory);
            }

            Task<ConstraintCatalogLoadResult> loadTask =
                Task.Factory.StartNew(
                    () => service.Load(options, BaseDirectory),
                    CancellationToken.None,
                    TaskCreationOptions.DenyChildAttach,
                    TaskScheduler.Default);
            try
            {
                if (loadTask.Wait(NetworkSourceTimeout))
                {
                    return loadTask.Result;
                }
            }
            catch (AggregateException exception)
            {
                return FailedNetworkLoad(
                    mode,
                    exception.GetBaseException().Message);
            }

            loadTask.ContinueWith(
                completed =>
                {
                    AggregateException ignored = completed.Exception;
                },
                CancellationToken.None,
                TaskContinuationOptions.OnlyOnFaulted |
                TaskContinuationOptions.ExecuteSynchronously,
                TaskScheduler.Default);
            return FailedNetworkLoad(
                mode,
                "network source did not respond within " +
                NetworkSourceTimeout.TotalSeconds.ToString("0") +
                " seconds");
        }

        private static ConstraintCatalogLoadResult FailedNetworkLoad(
            ConstraintSourceMode mode,
            string reason)
        {
            var result = new ConstraintCatalogLoadResult();
            result.Errors.Add(
                (mode == ConstraintSourceMode.RefDb ? "RefDB" : "Excel") +
                " could not be loaded: " + reason + ".");
            return result;
        }

        private static bool IsUncPath(string configuredPath)
        {
            return !string.IsNullOrWhiteSpace(configuredPath) &&
                (configuredPath.StartsWith(
                     @"\\",
                     StringComparison.Ordinal) ||
                 configuredPath.StartsWith(
                     @"\\?\UNC\",
                     StringComparison.OrdinalIgnoreCase));
        }

        public string GetConstraintTemplatesDirectory()
        {
            return EnsureDirectoryWithFallback(
                ResolvePath(Paths.ConstraintTemplatesDirectory),
                ResolvePath("ConstraintTemplates"));
        }

        public string GetLogsDirectory()
        {
            return EnsureDirectoryWithFallback(
                ResolvePath(Paths.LogsDirectory),
                ResolvePath("Logs"));
        }

        public string GetReportsDirectory()
        {
            return EnsureDirectoryWithFallback(
                ResolvePath(Paths.ReportsDirectory),
                ResolvePath("Reports"));
        }

        public string GetExportsDirectory()
        {
            return EnsureDirectoryWithFallback(
                ResolvePath(Paths.CsvExportDirectory),
                ResolvePath("Exports"));
        }

        public string GetStateDirectory()
        {
            return EnsureDirectoryWithFallback(
                ResolvePath(Paths.StateDirectory),
                ResolvePath("State"));
        }

        public string GetUsageLogFile()
        {
            return EnsureFileParentWithFallback(
                ResolvePath(Paths.UsageLogFile),
                ResolvePath(Path.Combine("Logs", "ClearPlan_UserLog.csv")));
        }

        public string GetActivityLogFile()
        {
            return EnsureFileParentWithFallback(
                ResolvePath(Paths.ActivityLogFile),
                ResolvePath(Path.Combine("Logs", "ActivityLog.csv")));
        }

        public string GetVersionSeenUsersFile()
        {
            return EnsureFileParentWithFallback(
                ResolvePath(Paths.VersionSeenUsersFile),
                ResolvePath(Path.Combine("State", "seen_versions.csv")));
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

        public void Save()
        {
            ApplyDefaults();
            string ini = PathSettingsIni.Serialize(this);
            var iniValidation = new ClearPlanSettings();
            PathSettingsIni.Apply(iniValidation, ini);
            string json = SerializeJsonOptions();
            JsonConvert.DeserializeObject<ClearPlanSettings>(json);
            WriteAtomically(PathSettingsFilePath, ini);
            WriteAtomically(SettingsFilePath, json);
            Reload();
        }

        private static ClearPlanSettings LoadInternal()
        {
            var settings = new ClearPlanSettings();
            settings.ApplyDefaults();
            string settingsJsonPath = Path.Combine(
                AssemblyHelper.GetAssemblyDirectory(),
                SettingsJsonFileName);
            if (File.Exists(settingsJsonPath))
            {
                try
                {
                    settings = JsonConvert.DeserializeObject<ClearPlanSettings>(
                        File.ReadAllText(settingsJsonPath, Encoding.UTF8)) ??
                        new ClearPlanSettings();
                }
                catch
                {
                    settings = new ClearPlanSettings();
                }
            }
            settings.ApplyDefaults();
            string iniPath = Path.Combine(
                AssemblyHelper.GetAssemblyDirectory(),
                PathSettingsIniFileName);
            if (File.Exists(iniPath))
            {
                try
                {
                    PathSettingsIni.Apply(
                        settings,
                        File.ReadAllText(iniPath, Encoding.UTF8));
                }
                catch
                {
                }
            }

            settings.ApplyDefaults();
            return settings;
        }

        private string SerializeJsonOptions()
        {
            var options = new
            {
                _comment =
                    "ClearPlan options. All file-system paths are stored in settings.ini.",
                constraintSource = new
                {
                    mode = ConstraintSource.Mode.ToString(),
                    includeInactiveTables = ConstraintSource.IncludeInactiveTables
                },
                links = Links,
                checks = Checks,
                privacy = Privacy
            };
            return JsonConvert.SerializeObject(options, Formatting.Indented);
        }

        private static void WriteAtomically(string target, string content)
        {
            string temporary = target + ".tmp." + Guid.NewGuid().ToString("N");
            File.WriteAllText(temporary, content, new UTF8Encoding(false));
            try
            {
                if (File.Exists(target))
                {
                    File.Replace(temporary, target, target + ".bak", true);
                }
                else
                {
                    File.Move(temporary, target);
                }
            }
            finally
            {
                if (File.Exists(temporary))
                {
                    File.Delete(temporary);
                }
            }
        }

        private void ApplyDefaults()
        {
            ConstraintSource = ConstraintSource ?? new ConstraintSourceOptions();
            Paths = Paths ?? new ClearPlanPathOptions();
            Links = Links ?? new ClearPlanLinkOptions();
            Checks = Checks ?? new ClearPlanCheckOptions();
            Privacy = Privacy ?? new ClearPlanPrivacyOptions();
            ClinicalPaths = ClinicalPaths ?? new ClearPlanClinicalPathOptions();

            ConstraintSource.ExcelWorkbookPath = DefaultIfBlank(
                ConstraintSource.ExcelWorkbookPath,
                Path.Combine("ConstraintTemplates", "ClearPlan_StockConstraints2024.xlsx"));
            ConstraintSource.RefDbJsonPath = ConstraintSource.RefDbJsonPath ?? string.Empty;
            ConstraintSource.StructureAliasesJsonPath = ConstraintSource.StructureAliasesJsonPath ?? string.Empty;
            Paths.MlcGeometryProfilesJsonPath = DefaultIfBlank(Paths.MlcGeometryProfilesJsonPath,
                Path.Combine("MachineGeometry", "MlcGeometryProfiles.example.json"));
            // An explicitly blank path disables dose-rate estimation; older settings receive the example.
            Paths.DoseRateProfilesJsonPath = Paths.DoseRateProfilesJsonPath ??
                Path.Combine("MachineGeometry", "DoseRateProfiles.example.json");
            Paths.AriaUploadConfigJsonPath = Paths.AriaUploadConfigJsonPath ?? string.Empty;
            // Null means an older configuration. An explicitly blank path disables Default rules.
            Paths.DefaultReviewRulesJsonPath = Paths.DefaultReviewRulesJsonPath ?? "DefaultReviewRules.json";
            Paths.FieldNamingRulesJsonPath = Paths.FieldNamingRulesJsonPath ?? "FieldNamingRules.json";

            Paths.ConstraintTemplatesDirectory = DefaultIfBlank(
                Paths.ConstraintTemplatesDirectory,
                "ConstraintTemplates");
            Paths.DefaultConventionalTemplate = DefaultIfBlank(
                Paths.DefaultConventionalTemplate,
                "ClearPlan_StockConstraints2024.xlsx");
            Paths.DefaultHypofractionatedTemplate = DefaultIfBlank(
                Paths.DefaultHypofractionatedTemplate,
                "ClearPlan_StockConstraints2024.xlsx");
            Paths.DefaultPlanSumTemplate = DefaultIfBlank(
                Paths.DefaultPlanSumTemplate,
                "ClearPlan_DefaultConstraints.xlsx");
            Paths.LogsDirectory = DefaultIfBlank(Paths.LogsDirectory, "Logs");
            Paths.ReportsDirectory = DefaultIfBlank(Paths.ReportsDirectory, "Reports");
            Paths.CsvExportDirectory = DefaultIfBlank(Paths.CsvExportDirectory, "Exports");
            Paths.StateDirectory = DefaultIfBlank(Paths.StateDirectory, "State");
            Paths.ConfigurationDirectory = DefaultIfBlank(Paths.ConfigurationDirectory, "Configuration");
            Paths.UsageLogFile = DefaultIfBlank(
                Paths.UsageLogFile,
                Path.Combine("Logs", "ClearPlan_UserLog.csv"));
            Paths.ActivityLogFile = DefaultIfBlank(
                Paths.ActivityLogFile,
                Path.Combine("Logs", "ActivityLog.csv"));
            Paths.VersionSeenUsersFile = DefaultIfBlank(
                Paths.VersionSeenUsersFile,
                Path.Combine("State", "seen_versions.csv"));
            Paths.ChangeLogFile = DefaultIfBlank(Paths.ChangeLogFile, "CHANGELOG.md");
            Paths.FeedbackFile = DefaultIfBlank(Paths.FeedbackFile, "FEEDBACK.md");

            Links.FeedbackUrl = Links.FeedbackUrl ?? string.Empty;
            Checks.Profile = DefaultIfBlank(Checks.Profile, "starter");
            Privacy.HashSalt = Privacy.HashSalt ?? string.Empty;
            ClinicalPaths.PrescriptionSettingsPath =
                ClinicalPaths.PrescriptionSettingsPath ?? string.Empty;
            ClinicalPaths.PlanCheckResourcesPath =
                ClinicalPaths.PlanCheckResourcesPath ?? string.Empty;
        }

        private static string DefaultIfBlank(string value, string fallbackValue)
        {
            return string.IsNullOrWhiteSpace(value) ? fallbackValue : value;
        }

        private static string EnsureDirectoryWithFallback(
            string configuredPath,
            string fallbackPath)
        {
            WritablePathResult result = WritablePathFallback.EnsureDirectory(
                configuredPath,
                fallbackPath);
            TraceFallback(configuredPath, result);
            return result.Path;
        }

        private static string EnsureFileParentWithFallback(
            string configuredPath,
            string fallbackPath)
        {
            WritablePathResult result = WritablePathFallback.EnsureFileParent(
                configuredPath,
                fallbackPath);
            TraceFallback(configuredPath, result);
            return result.Path;
        }

        private static void TraceFallback(
            string configuredPath,
            WritablePathResult result)
        {
            if (result.UsedFallback)
            {
                System.Diagnostics.Trace.TraceWarning(
                    "ClearPlan path unavailable. Configured='{0}', fallback='{1}', reason='{2}'",
                    configuredPath,
                    result.Path,
                    result.FailureMessage);
            }
        }
    }
}
