using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Review;
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
            StructureAliases = new ObservableCollection<StructureAliasRowViewModel>();
            PlanCheckSelectionRows = new ObservableCollection<PlanCheckSelectionRowViewModel>();
        }

        public IEnumerable<ConstraintSourceMode> SourceModes { get; private set; }
        public ConstraintSourceMode SourceMode { get; set; }
        public string RefDbJsonPath { get; set; }
        public string ExcelWorkbookPath { get; set; }
        public string StructureAliasesJsonPath { get; set; }
        public string MlcGeometryProfilesJsonPath { get; set; }
        public string DoseRateProfilesJsonPath { get; set; }
        public string CollisionProfilesJsonPath { get; set; }
        public string SourceCollisionModelsJsonPath { get; set; }
        public string AriaUploadConfigJsonPath { get; set; }
        public string DefaultReviewRulesJsonPath { get; set; }
        public string PlanCheckSelectionJsonPath { get; set; }
        public string FieldNamingRulesJsonPath { get; set; }
        public string IsodoseDisplayJsonPath { get; set; }
        public string ConfigurationDirectory { get; set; }
        public ConfigurationWorkspaceViewModel Configuration { get; private set; }
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
        public ObservableCollection<StructureAliasRowViewModel> StructureAliases { get; private set; }
        public ObservableCollection<PlanCheckSelectionRowViewModel> PlanCheckSelectionRows { get; private set; }
        public string PlanCheckSelectionStatus { get; private set; }
        public bool PlanCheckSelectionDirty { get; private set; }
        private ClearPlanSettings _configurationSettings;

        public void InitializeCheckSelection(IEnumerable<ReviewCheckRow> rows)
        {
            string path = string.IsNullOrWhiteSpace(PlanCheckSelectionJsonPath) ? null :
                _configurationSettings == null ? PlanCheckSelectionJsonPath : _configurationSettings.ResolvePath(PlanCheckSelectionJsonPath);
            var config = PlanCheckSelectionConfiguration.Load(path);
            var groups = (rows ?? Enumerable.Empty<ReviewCheckRow>()).Where(PlanCheckSelectionConfiguration.CanConfigure)
                .GroupBy(row => PlanCheckSelectionConfiguration.BaseCheckCode(row.CheckCode), StringComparer.Ordinal)
                .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);
            PlanCheckSelectionRows.Clear();
            foreach (string code in groups.Keys.Concat(config.Checks.Select(check => check.Code)).Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal))
            {
                List<ReviewCheckRow> findings;
                groups.TryGetValue(code, out findings);
                PlanCheckSelectionRows.Add(new PlanCheckSelectionRowViewModel {
                    Code = code, Enabled = config.IsEnabled(code), FindingCount = findings == null ? 0 : findings.Count,
                    Description = findings == null ? "Konfigurierte Check-Familie; kein Befund im übergebenen Review." :
                        string.Join(" · ", findings.Select(row => row.Message).Where(message => !string.IsNullOrWhiteSpace(message)).Distinct().Take(3)),
                    Changed = () => { PlanCheckSelectionDirty = true; NotifyPropertyChanged("PlanCheckSelectionDirty"); } });
            }
            PlanCheckSelectionDirty = false;
            PlanCheckSelectionStatus = config.Message + " Übersicht: " + PlanCheckSelectionRows.Count + " Check-Familien. Beschreibungen bleiben nur im Arbeitsspeicher.";
            NotifyPropertyChanged("PlanCheckSelectionStatus");
            NotifyPropertyChanged("PlanCheckSelectionDirty");
        }

        public void StageCheckSelection()
        {
            if (Configuration == null) throw new InvalidOperationException("Versionsablage nicht zugänglich.");
            // Never serialize the row view models: descriptions can contain current clinical context.
            string json = PlanCheckSelectionConfiguration.Serialize(PlanCheckSelectionRows.Select(row =>
                new PlanCheckSelectionEntry { Code = row.Code, Enabled = row.Enabled }));
            Configuration.StageJson("plancheck-selection", json);
            PlanCheckSelectionDirty = false;
            NotifyPropertyChanged("PlanCheckSelectionDirty");
        }

        public void SetAliases(IEnumerable<StructureDefinition> definitions)
        {
            StructureAliases.Clear();
            foreach (StructureDefinition row in definitions ?? Enumerable.Empty<StructureDefinition>())
                StructureAliases.Add(StructureAliasRowViewModel.From(row));
        }

        public IList<StructureDefinition> GetAliases()
        {
            return StructureAliases.Select(row => row.ToDefinition()).ToList();
        }

        public static SettingsViewModel From(
            ClearPlanSettings settings,
            ConstraintCatalogLoadResult loadResult)
        {
            return new SettingsViewModel
            {
                SourceMode = settings.ConstraintSource.Mode,
                RefDbJsonPath = settings.ConstraintSource.RefDbJsonPath,
                ExcelWorkbookPath = settings.ConstraintSource.ExcelWorkbookPath,
                StructureAliasesJsonPath = settings.ConstraintSource.StructureAliasesJsonPath,
                MlcGeometryProfilesJsonPath = settings.Paths.MlcGeometryProfilesJsonPath,
                DoseRateProfilesJsonPath = settings.Paths.DoseRateProfilesJsonPath,
                CollisionProfilesJsonPath = settings.Paths.CollisionProfilesJsonPath,
                SourceCollisionModelsJsonPath = settings.Paths.SourceCollisionModelsJsonPath,
                AriaUploadConfigJsonPath = settings.Paths.AriaUploadConfigJsonPath,
                DefaultReviewRulesJsonPath = settings.Paths.DefaultReviewRulesJsonPath,
                PlanCheckSelectionJsonPath = settings.Paths.PlanCheckSelectionJsonPath,
                FieldNamingRulesJsonPath = settings.Paths.FieldNamingRulesJsonPath,
                IsodoseDisplayJsonPath = settings.Paths.IsodoseDisplayJsonPath,
                ConfigurationDirectory = settings.Paths.ConfigurationDirectory,
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
            settings.Paths.MlcGeometryProfilesJsonPath = MlcGeometryProfilesJsonPath ?? string.Empty;
            settings.Paths.DoseRateProfilesJsonPath = DoseRateProfilesJsonPath ?? string.Empty;
            settings.Paths.CollisionProfilesJsonPath = CollisionProfilesJsonPath ?? string.Empty;
            settings.Paths.SourceCollisionModelsJsonPath = SourceCollisionModelsJsonPath ?? string.Empty;
            settings.Paths.AriaUploadConfigJsonPath = AriaUploadConfigJsonPath ?? string.Empty;
            settings.Paths.DefaultReviewRulesJsonPath = DefaultReviewRulesJsonPath ?? string.Empty;
            settings.Paths.PlanCheckSelectionJsonPath = PlanCheckSelectionJsonPath ?? string.Empty;
            settings.Paths.FieldNamingRulesJsonPath = FieldNamingRulesJsonPath ?? string.Empty;
            settings.Paths.IsodoseDisplayJsonPath = IsodoseDisplayJsonPath ?? string.Empty;
            settings.Paths.ConfigurationDirectory = ConfigurationDirectory ?? "Configuration";
            settings.ConstraintSource.RefDbJsonPath =
                RefDbJsonPath ?? string.Empty;
            settings.ConstraintSource.ExcelWorkbookPath =
                ExcelWorkbookPath ?? string.Empty;
            settings.ConstraintSource.StructureAliasesJsonPath =
                StructureAliasesJsonPath ?? string.Empty;
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

        public void InitializeConfiguration(ClearPlanSettings settings)
        {
            _configurationSettings = settings;
            string root = settings.ResolvePath(string.IsNullOrWhiteSpace(ConfigurationDirectory) ? "Configuration" : ConfigurationDirectory);
            var entries = new[]
            {
                Entry("default-rules", "Default-Zielregeln", "Zieltypen, DVH-Metriken und Verordnungsanteile. Deaktivierte oder entfernte Regeln werden nicht ersetzt.", ".json", () => DefaultReviewRulesJsonPath),
                Entry("isodose-display", "Default · Isodosen", "Anzeige in Schnittbildern und Reports: Prozent der Plan-Gesamtverordnung, Farbe #RRGGBB und Aktivierung. Nur Darstellung, keine Änderung der Plandosis oder Goals. Änderungen und Rückkehr zu älteren Versionen werden protokolliert.", ".json", () => IsodoseDisplayJsonPath),
                Entry("plancheck-selection", "PlanCheck · Review-Auswahl", "Einbeziehen bestehender PlanCheck-Befunde in Review und Bericht. Die externe Berechnung läuft weiterhin; native Eclipse-Warnungen bleiben sichtbar. Nur Check-Codes und Aktivierung werden gespeichert.", ".json", () => PlanCheckSelectionJsonPath),
                Entry("constraints", "Constraints · Excel", "ClearPlan-Katalog mit Tables, Constraints und Structures. Excel-Entwürfe werden erst nach ausdrücklicher Prüfung übernommen.", ".xlsx", () => ExcelWorkbookPath),
                Entry("aliases", "Strukturnamen & Aliase", "Explizite alternative Strukturnamen. Keine Umbenennung von Strukturen im Planungssystem.", ".json", () => StructureAliasesJsonPath),
                Entry("field-naming", "Default · Feldnamen", "Nomenklatur, z. B. 120-30 T300 GUZ: Tisch vorletzter Bestandteil, Richtung letzter. Reihenfolge, Trennzeichen und Richtungstoken sind konfigurierbar. Nur Namensvorschau; keine Änderung an Bestrahlungsfeldern.", ".json", () => FieldNamingRulesJsonPath),
                Entry("mlc-profiles", "MLC-Geometrieprofile", "Exakte Modellnamen, Blattgrenzen und bestätigte Lagenzuordnung. Keine klinische Kommissionierung durch die Dateiprüfung.", ".json", () => MlcGeometryProfilesJsonPath),
                Entry("collision-profiles", "Kollision · Gerätehüllen", "Vermessene Kopf-/Bore-Hüllen mit exakter Geräte-ID, Revision und Kommissionierungsnachweis. Keine Geräteabmessungen aus MLC-Typen ableiten. Fehlende Profile bleiben nicht bewertbar; Versionen sind wiederherstellbar.", ".json", () => CollisionProfilesJsonPath),
                Entry("collision-source-models", "Kollision · Quellmodelle", "Analytische Quellmodelle und exakte lokale Gerätezuordnungen für ein Punkt-Screening. Herkunft und Messangaben werden dokumentiert; keine klinische Freigabe, keine kommissionierte Abstandsberechnung.", ".json", () => SourceCollisionModelsJsonPath),
                Entry("dose-rate-profiles", "Dosisrate · Schätzmodell", "PlanCheck-Schätzung mit expliziten Maschinenzuordnungen und angenommener Gantry-Geschwindigkeit. Keine gemessene Abgabe; Beschleunigung, MLC-/Blendenbewegung und Beam-Holds sind nicht modelliert. Versionen können geprüft und wiederhergestellt werden.", ".json", () => DoseRateProfilesJsonPath),
                Entry("aria-upload", "ARIA · Berichtversand", "Lokale HTTPS-Endpunkte, Dokumenttyp und Verweis auf externe Zugangsdaten. Keine Secrets im JSON. Jeder Upload wird bestätigt; kein Schreiben am Bestrahlungsplan.", ".json", () => AriaUploadConfigJsonPath),
                Entry("refdb", "RefDB · optionale Kopie", "Externe RefDB bleibt unverändert. Ein ausdrücklicher Import erzeugt eine getrennte Momentaufnahme; spätere RefDB-Aktualisierungen werden nicht automatisch übernommen.", ".json", () => RefDbJsonPath)
            };
            foreach (var entry in entries)
            {
                var captured = entry;
                var source = captured.SourcePath;
                captured.SourcePath = () =>
                {
                    string path = source();
                    return string.IsNullOrWhiteSpace(path) ? string.Empty : settings.ResolvePath(path);
                };
                captured.Validate = bytes => ConfigurationContentValidator.Validate(captured.Key, bytes, root);
            }
            Configuration = new ConfigurationWorkspaceViewModel(root, entries, (key, path) =>
            {
                if (key == "default-rules") DefaultReviewRulesJsonPath = path;
                else if (key == "isodose-display") IsodoseDisplayJsonPath = path;
                else if (key == "plancheck-selection") PlanCheckSelectionJsonPath = path;
                else if (key == "constraints") ExcelWorkbookPath = path;
                else if (key == "aliases") StructureAliasesJsonPath = path;
                else if (key == "field-naming") FieldNamingRulesJsonPath = path;
                else if (key == "mlc-profiles") MlcGeometryProfilesJsonPath = path;
                else if (key == "collision-profiles") CollisionProfilesJsonPath = path;
                else if (key == "collision-source-models") SourceCollisionModelsJsonPath = path;
                else if (key == "dose-rate-profiles") DoseRateProfilesJsonPath = path;
                else if (key == "aria-upload") AriaUploadConfigJsonPath = path;
                else if (key == "refdb") RefDbJsonPath = path;
                NotifyPropertyChanged(string.Empty);
            });
            NotifyPropertyChanged("Configuration");
        }

        private static ConfigurationEntryViewModel Entry(string key, string title, string description,
            string extension, Func<string> sourcePath)
        {
            return new ConfigurationEntryViewModel { Key = key, Title = title, Description = description,
                Extension = extension, SourcePath = sourcePath };
        }
    }

    public sealed class PlanCheckSelectionRowViewModel : ViewModelBase
    {
        private bool _enabled;
        public string Code { get; set; }
        public string Description { get; set; }
        public int FindingCount { get; set; }
        public Action Changed { get; set; }
        public bool Enabled
        {
            get { return _enabled; }
            set { if (_enabled == value) return; _enabled = value; NotifyPropertyChanged("Enabled"); if (Changed != null) Changed(); }
        }
    }

    public sealed class StructureAliasRowViewModel
    {
        public string StructureId { get; set; }
        public string CanonicalName { get; set; }
        public string Aliases { get; set; }
        public string Laterality { get; set; }
        public string SideAliasesLeft { get; set; }
        public string SideAliasesRight { get; set; }

        public static StructureAliasRowViewModel From(StructureDefinition row)
        {
            return new StructureAliasRowViewModel
            {
                StructureId = row.StructureId, CanonicalName = row.CanonicalName, Laterality = row.Laterality,
                Aliases = Join(row.Aliases), SideAliasesLeft = Join(row.SideAliasesLeft), SideAliasesRight = Join(row.SideAliasesRight)
            };
        }

        public StructureDefinition ToDefinition()
        {
            return new StructureDefinition
            {
                StructureId = StructureId, CanonicalName = (CanonicalName ?? string.Empty).Trim(), Active = true,
                Laterality = (Laterality ?? string.Empty).Trim(), Aliases = Split(Aliases),
                SideAliasesLeft = Split(SideAliasesLeft), SideAliasesRight = Split(SideAliasesRight)
            };
        }

        private static string Join(IEnumerable<string> values) { return string.Join(" | ", values ?? Enumerable.Empty<string>()); }
        private static IList<string> Split(string value)
        {
            return (value ?? string.Empty).Split(new[] { '|', ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(item => item.Trim()).Where(item => item.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }
    }
}
