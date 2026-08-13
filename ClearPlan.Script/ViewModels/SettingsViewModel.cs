using System;
using System.Collections.Generic;
using ClearPlan.Core.Constraints;
using ClearPlan.Helpers;

namespace ClearPlan
{
    public sealed class SettingsViewModel : ViewModelBase
    {
        public SettingsViewModel()
        {
            SourceModes = new[]
            {
                ConstraintSourceMode.Automatic,
                ConstraintSourceMode.RefDb,
                ConstraintSourceMode.Excel
            };
        }

        public IEnumerable<ConstraintSourceMode> SourceModes { get; private set; }
        public ConstraintSourceMode SourceMode { get; set; }
        public string RefDbJsonPath { get; set; }
        public string ExcelWorkbookPath { get; set; }
        public bool IncludeInactiveTables { get; set; }
        public string ConstraintTemplatesDirectory { get; set; }
        public string DefaultConventionalTemplate { get; set; }
        public string DefaultHypofractionatedTemplate { get; set; }
        public string DefaultPlanSumTemplate { get; set; }
        public string LogsDirectory { get; set; }
        public string ReportsDirectory { get; set; }
        public string ExportsDirectory { get; set; }
        public string StateDirectory { get; set; }
        public string UsageLogFile { get; set; }
        public string ActivityLogFile { get; set; }
        public string VersionSeenUsersFile { get; set; }
        public string ChangeLogFile { get; set; }
        public string FeedbackFile { get; set; }
        public string PrescriptionSettingsPath { get; set; }
        public string PlanCheckResourcesPath { get; set; }
        public bool LogPatientIdentifiers { get; set; }
        public bool LogUserIdentifiers { get; set; }
        public bool UsePatientIdentifiersInFilenames { get; set; }
        public ConstraintSourceStatusViewModel SourceStatus { get; set; }

        public static SettingsViewModel From(
            ClearPlanSettings settings,
            ConstraintCatalogLoadResult loadResult)
        {
            return new SettingsViewModel
            {
                SourceMode = settings.ConstraintSource.Mode,
                RefDbJsonPath = settings.ConstraintSource.RefDbJsonPath,
                ExcelWorkbookPath = settings.ConstraintSource.ExcelWorkbookPath,
                IncludeInactiveTables = settings.ConstraintSource.IncludeInactiveTables,
                ConstraintTemplatesDirectory =
                    settings.Paths.ConstraintTemplatesDirectory,
                DefaultConventionalTemplate =
                    settings.Paths.DefaultConventionalTemplate,
                DefaultHypofractionatedTemplate =
                    settings.Paths.DefaultHypofractionatedTemplate,
                DefaultPlanSumTemplate =
                    settings.Paths.DefaultPlanSumTemplate,
                LogsDirectory = settings.Paths.LogsDirectory,
                ReportsDirectory = settings.Paths.ReportsDirectory,
                ExportsDirectory = settings.Paths.CsvExportDirectory,
                StateDirectory = settings.Paths.StateDirectory,
                UsageLogFile = settings.Paths.UsageLogFile,
                ActivityLogFile = settings.Paths.ActivityLogFile,
                VersionSeenUsersFile = settings.Paths.VersionSeenUsersFile,
                ChangeLogFile = settings.Paths.ChangeLogFile,
                FeedbackFile = settings.Paths.FeedbackFile,
                PrescriptionSettingsPath =
                    settings.ClinicalPaths.PrescriptionSettingsPath,
                PlanCheckResourcesPath =
                    settings.ClinicalPaths.PlanCheckResourcesPath,
                LogPatientIdentifiers =
                    settings.Privacy.LogPatientIdentifiers,
                LogUserIdentifiers =
                    settings.Privacy.LogUserIdentifiers,
                UsePatientIdentifiersInFilenames =
                    settings.Privacy.UsePatientIdentifiersInFilenames,
                SourceStatus = ConstraintSourceStatusViewModel.From(
                    loadResult,
                    !string.IsNullOrWhiteSpace(
                        settings.ConstraintSource.RefDbJsonPath))
            };
        }

        public void ApplyTo(ClearPlanSettings settings)
        {
            settings.ConstraintSource.Mode = SourceMode;
            settings.ConstraintSource.RefDbJsonPath =
                RefDbJsonPath ?? string.Empty;
            settings.ConstraintSource.ExcelWorkbookPath =
                ExcelWorkbookPath ?? string.Empty;
            settings.ConstraintSource.IncludeInactiveTables =
                IncludeInactiveTables;
            settings.Paths.ConstraintTemplatesDirectory =
                ConstraintTemplatesDirectory ?? string.Empty;
            settings.Paths.DefaultConventionalTemplate =
                DefaultConventionalTemplate ?? string.Empty;
            settings.Paths.DefaultHypofractionatedTemplate =
                DefaultHypofractionatedTemplate ?? string.Empty;
            settings.Paths.DefaultPlanSumTemplate =
                DefaultPlanSumTemplate ?? string.Empty;
            settings.Paths.LogsDirectory = LogsDirectory ?? string.Empty;
            settings.Paths.ReportsDirectory = ReportsDirectory ?? string.Empty;
            settings.Paths.CsvExportDirectory = ExportsDirectory ?? string.Empty;
            settings.Paths.StateDirectory = StateDirectory ?? string.Empty;
            settings.Paths.UsageLogFile = UsageLogFile ?? string.Empty;
            settings.Paths.ActivityLogFile = ActivityLogFile ?? string.Empty;
            settings.Paths.VersionSeenUsersFile =
                VersionSeenUsersFile ?? string.Empty;
            settings.Paths.ChangeLogFile = ChangeLogFile ?? string.Empty;
            settings.Paths.FeedbackFile = FeedbackFile ?? string.Empty;
            settings.ClinicalPaths.PrescriptionSettingsPath =
                PrescriptionSettingsPath ?? string.Empty;
            settings.ClinicalPaths.PlanCheckResourcesPath =
                PlanCheckResourcesPath ?? string.Empty;
            settings.Privacy.LogPatientIdentifiers = LogPatientIdentifiers;
            settings.Privacy.LogUserIdentifiers = LogUserIdentifiers;
            settings.Privacy.UsePatientIdentifiersInFilenames =
                UsePatientIdentifiersInFilenames;
        }
    }
}
