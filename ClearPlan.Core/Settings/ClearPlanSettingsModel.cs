using ClearPlan.Core.Constraints;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace ClearPlan.Core.Settings
{
    public class ClearPlanSettingsModel
    {
        public ClearPlanSettingsModel()
        {
            ConstraintSource = new ConstraintSourceOptions();
            Paths = new ClearPlanPathOptions();
            Links = new ClearPlanLinkOptions();
            Checks = new ClearPlanCheckOptions();
            Privacy = new ClearPlanPrivacyOptions();
            ClinicalPaths = new ClearPlanClinicalPathOptions();
        }

        public ConstraintSourceOptions ConstraintSource { get; set; }
        public ClearPlanPathOptions Paths { get; set; }
        public ClearPlanLinkOptions Links { get; set; }
        public ClearPlanCheckOptions Checks { get; set; }
        public ClearPlanPrivacyOptions Privacy { get; set; }
        public ClearPlanClinicalPathOptions ClinicalPaths { get; set; }
    }

    public sealed class ConstraintSourceOptions
    {
        [JsonConverter(typeof(StringEnumConverter))]
        public ConstraintSourceMode Mode { get; set; } = ConstraintSourceMode.Automatic;

        public string RefDbJsonPath { get; set; }
        public string ExcelWorkbookPath { get; set; }
        public string StructureAliasesJsonPath { get; set; }
        public bool IncludeInactiveTables { get; set; }
    }

    public sealed class ClearPlanPathOptions
    {
        public string MlcGeometryProfilesJsonPath { get; set; }
        public string DoseRateProfilesJsonPath { get; set; }
        public string CollisionProfilesJsonPath { get; set; }
        public string SourceCollisionModelsJsonPath { get; set; }
        public string AriaUploadConfigJsonPath { get; set; }
        public string DefaultReviewRulesJsonPath { get; set; }
        public string PlanCheckSelectionJsonPath { get; set; }
        public string FieldNamingRulesJsonPath { get; set; }
        public string IsodoseDisplayJsonPath { get; set; }
        public string ConstraintTemplatesDirectory { get; set; }
        public string DefaultConventionalTemplate { get; set; }
        public string DefaultHypofractionatedTemplate { get; set; }
        public string DefaultPlanSumTemplate { get; set; }
        public string LogsDirectory { get; set; }
        public string ReportsDirectory { get; set; }
        public string CsvExportDirectory { get; set; }
        public string StateDirectory { get; set; }
        public string ConfigurationDirectory { get; set; }
        public string UsageLogFile { get; set; }
        public string ActivityLogFile { get; set; }
        public string VersionSeenUsersFile { get; set; }
        public string ChangeLogFile { get; set; }
        public string FeedbackFile { get; set; }
    }

    public sealed class ClearPlanLinkOptions
    {
        public string FeedbackUrl { get; set; }
    }

    public sealed class ClearPlanCheckOptions
    {
        public string Profile { get; set; }
    }

    public sealed class ClearPlanPrivacyOptions
    {
        public bool LogPatientIdentifiers { get; set; }
        public bool LogUserIdentifiers { get; set; }
        public bool UsePatientIdentifiersInFilenames { get; set; }
        public string HashSalt { get; set; }
    }

    public sealed class ClearPlanClinicalPathOptions
    {
        public string PrescriptionSettingsPath { get; set; }
        public string PlanCheckResourcesPath { get; set; }
    }
}
