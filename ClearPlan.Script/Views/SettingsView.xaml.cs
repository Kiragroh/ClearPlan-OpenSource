using System;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ClearPlan.Core.Configuration;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Review;
using System.Collections.Generic;
using ClearPlan.Core.Settings;
using ClearPlan.Helpers;
using Microsoft.Win32;
using Newtonsoft.Json;

namespace ClearPlan.Views
{
    public partial class SettingsView : UserControl
    {
        private SettingsViewModel _viewModel;
        private bool _aliasesDirty;
        private List<ReviewCheckRow> _checkSelectionRows = new List<ReviewCheckRow>();

        public SettingsView()
        {
            InitializeComponent();
            ReloadView();
        }

        public event EventHandler SettingsChanged;

        public void InitializeCheckSelection(IEnumerable<ReviewCheckRow> rows)
        {
            _checkSelectionRows = (rows ?? Enumerable.Empty<ReviewCheckRow>()).ToList();
            if (_viewModel != null && !_viewModel.PlanCheckSelectionDirty)
                _viewModel.InitializeCheckSelection(_checkSelectionRows);
        }

        private void LoadCheckSelectionClicked(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscardCheckSelection()) return;
            _viewModel.InitializeCheckSelection(_checkSelectionRows);
        }

        private void StageCheckSelectionClicked(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscardConfiguration()) return;
            CheckSelectionGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            CheckSelectionGrid.CommitEdit(DataGridEditingUnit.Row, true);
            try { _viewModel.StageCheckSelection(); SettingsTabs.SelectedIndex = 0; }
            catch (Exception exception) { CheckSelectionStatusText.Text = "Entwurf nicht übernommen: " + exception.Message; }
        }

        private bool ConfirmDiscardCheckSelection()
        {
            return _viewModel == null || !_viewModel.PlanCheckSelectionDirty || MessageBox.Show(
                "Ungespeicherte PlanCheck-Auswahl verwerfen?", "PlanCheck-Auswahl", MessageBoxButton.YesNo,
                MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes;
        }

        private void BrowseDefaultRulesClicked(object sender, RoutedEventArgs e)
        {
            string path = BrowseFile("Default-Regeln wählen", "Default-Regeln (*.json)|*.json");
            if (path == null) return;
            _viewModel.DefaultReviewRulesJsonPath = path;
            RefreshBindings();
            DefaultRulesStatusText.Text = "Datei gewählt. Zum Aktivieren Einstellungen speichern und Planansicht neu laden.";
        }

        private void EditDefaultRulesClicked(object sender, RoutedEventArgs e)
        {
            if (_viewModel.Configuration == null)
            {
                DefaultRulesStatusText.Text = "Versionsablage nicht zugänglich. Den Ablagepfad oben korrigieren und Einstellungen speichern.";
                return;
            }
            if (!ConfirmDiscardConfiguration()) return;
            _viewModel.Configuration.Select(_viewModel.Configuration.Entries.Single(entry => entry.Key == "default-rules"));
            SettingsTabs.SelectedIndex = 0;
        }

        private void BrowseAliasesClicked(object sender, RoutedEventArgs e)
        {
            string path = BrowseFile("Separate Alias-JSON öffnen", "Alias-JSON (*.json)|*.json");
            if (path == null || !ConfirmDiscardAliases()) return;
            try
            {
                _viewModel.SetAliases(StructureAliasConfiguration.Deserialize(File.ReadAllText(path, System.Text.Encoding.UTF8)));
                _viewModel.StructureAliasesJsonPath = path;
                _aliasesDirty = false;
                RefreshBindings();
                AliasStatusText.Text = "Aliasdatei geöffnet. Zum Aktivieren Einstellungen speichern.";
            }
            catch (Exception exception) { AliasStatusText.Text = "Nicht geöffnet: " + exception.Message; }
        }

        private void ImportAliasesClicked(object sender, RoutedEventArgs e)
        {
            string path = BrowseFile("Strukturzuordnungen importieren", "Alias-/RefDB-JSON oder Constraint-Excel (*.json;*.xlsx)|*.json;*.xlsx");
            if (path == null || !ConfirmDiscardAliases()) return;
            try
            {
                _viewModel.SetAliases(StructureAliasConfiguration.Import(path));
                _aliasesDirty = true;
                AliasStatusText.Text = "Importiert, noch nicht gespeichert. Zuordnungen prüfen und als separate Alias-JSON speichern. Quelldatei bleibt unverändert.";
            }
            catch (Exception exception) { AliasStatusText.Text = "Import nicht möglich: " + exception.Message; }
        }

        private void LoadSourceAliasesClicked(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscardAliases()) return;
            ClearPlanSettings source = CreateDraft();
            source.ConstraintSource.StructureAliasesJsonPath = string.Empty;
            ConstraintCatalogLoadResult catalog = source.LoadConstraintCatalog();
            if (!catalog.IsUsable)
            {
                AliasStatusText.Text = string.Join(Environment.NewLine, catalog.Errors);
                return;
            }
            _viewModel.SetAliases(catalog.Catalog.Structures);
            _aliasesDirty = true;
            AliasStatusText.Text = "Strukturzuordnungen aus der gewählten Quelle übernommen. Änderungen sind noch nicht gespeichert.";
        }

        private void AddAliasClicked(object sender, RoutedEventArgs e)
        {
            var row = new StructureAliasRowViewModel();
            _viewModel.StructureAliases.Add(row);
            AliasesGrid.SelectedItem = row;
            AliasesGrid.ScrollIntoView(row);
            _aliasesDirty = true;
        }

        private void RemoveAliasClicked(object sender, RoutedEventArgs e)
        {
            foreach (StructureAliasRowViewModel row in AliasesGrid.SelectedItems.Cast<StructureAliasRowViewModel>().ToList())
                _viewModel.StructureAliases.Remove(row);
            _aliasesDirty = true;
            AliasStatusText.Text = "Auswahl aus dem Entwurf entfernt. Nicht aufgeführte Strukturen behalten die Zuordnung der Quelle.";
        }

        private void AliasRowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            if (e.EditAction == DataGridEditAction.Commit)
            {
                _aliasesDirty = true;
                AliasStatusText.Text = "Aliasänderungen noch nicht gespeichert.";
            }
        }

        private void ValidateAliasesClicked(object sender, RoutedEventArgs e) { ValidateAliasRows(); }

        private bool ValidateAliasRows()
        {
            AliasesGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            AliasesGrid.CommitEdit(DataGridEditingUnit.Row, true);
            var errors = StructureAliasConfiguration.Validate(_viewModel.GetAliases());
            if (errors.Count > 0)
            {
                AliasStatusText.Text = "Nicht gültig:" + Environment.NewLine + string.Join(Environment.NewLine, errors);
                return false;
            }
            try
            {
                ClearPlanSettings source = CreateDraft();
                source.ConstraintSource.StructureAliasesJsonPath = string.Empty;
                ConstraintCatalogLoadResult catalog = source.LoadConstraintCatalog();
                if (!catalog.IsUsable)
                {
                    AliasStatusText.Text = "Abgleich mit der Constraint-Quelle nicht möglich: " + string.Join(" ", catalog.Errors);
                    return false;
                }
                StructureAliasConfiguration.Merge(catalog.Catalog.Structures, _viewModel.GetAliases());
                AliasStatusText.Text = "Keine mehrdeutigen Zuordnungen. TG-263-Schreibweise und anatomische Gleichwertigkeit müssen fachlich geprüft werden.";
                return true;
            }
            catch (Exception exception) { AliasStatusText.Text = "Nicht gültig: " + exception.Message; return false; }
        }

        private void ExportAliasesClicked(object sender, RoutedEventArgs e) { WriteAliases(false); }
        private void SaveAliasesClicked(object sender, RoutedEventArgs e) { WriteAliases(true); }

        private void WriteAliases(bool useForSettings)
        {
            if (!ValidateAliasRows()) return;
            if (useForSettings)
            {
                if (_viewModel.Configuration == null)
                {
                    AliasStatusText.Text = "Versionsablage nicht zugänglich. Den Ablagepfad oben korrigieren und Einstellungen speichern.";
                    return;
                }
                if (!ConfirmDiscardConfiguration()) return;
                _viewModel.Configuration.StageJson("aliases", StructureAliasConfiguration.Serialize(_viewModel.GetAliases()));
                _aliasesDirty = false;
                SettingsTabs.SelectedIndex = 0;
                AliasStatusText.Text = "Aliasentwurf in die Konfigurationsverwaltung übertragen. Dort mit Änderungsgrund als Version speichern.";
                return;
            }
            var dialog = new SaveFileDialog
            {
                Title = useForSettings ? "Separate Aliasdatei speichern" : "Alias-JSON exportieren",
                Filter = "Alias-JSON (*.json)|*.json", DefaultExt = ".json", AddExtension = true,
                OverwritePrompt = true, FileName = "StructureAliases.json"
            };
            if (dialog.ShowDialog() != true) return;
            try
            {
                StructureAliasConfiguration.Save(dialog.FileName, _viewModel.GetAliases());
                if (useForSettings)
                {
                    _viewModel.StructureAliasesJsonPath = dialog.FileName;
                    _aliasesDirty = false;
                    RefreshBindings();
                }
                AliasStatusText.Text = useForSettings
                    ? "Aliasdatei gespeichert. Jetzt unten Einstellungen speichern, um die Datei zu aktivieren."
                    : "Aliasdatei exportiert; aktive Einstellungen unverändert.";
            }
            catch (Exception exception) { AliasStatusText.Text = "Nicht gespeichert: " + exception.Message; }
        }

        private bool ConfirmDiscardAliases()
        {
            return !_aliasesDirty || MessageBox.Show("Ungespeicherten Aliasentwurf verwerfen?", "Struktur-Aliase",
                MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes;
        }

        private void BrowseRefDbClicked(object sender, RoutedEventArgs e)
        {
            string path = BrowseFile(
                "RefDB JSON auswählen",
                "JSON-Datei (*.json)|*.json|Alle Dateien (*.*)|*.*");
            if (!string.IsNullOrWhiteSpace(path))
            {
                _viewModel.RefDbJsonPath = path;
                RefreshBindings();
            }
        }

        private void BrowseExcelClicked(object sender, RoutedEventArgs e)
        {
            string path = BrowseFile(
                "Constraint-Arbeitsmappe auswählen",
                "Excel-Arbeitsmappe (*.xlsx)|*.xlsx|Alle Dateien (*.*)|*.*");
            if (!string.IsNullOrWhiteSpace(path))
            {
                _viewModel.ExcelWorkbookPath = path;
                RefreshBindings();
            }
        }

        private void ValidateClicked(object sender, RoutedEventArgs e)
        {
            ClearPlanSettings draft = CreateDraft();
            SettingsValidationResult settingsResult =
                SettingsValidator.Validate(draft, draft.BaseDirectory);
            if (!settingsResult.IsValid)
            {
                ActionStatusText.Text = string.Join(
                    Environment.NewLine,
                    settingsResult.Errors);
                return;
            }

            ConstraintCatalogLoadResult catalogResult =
                draft.LoadConstraintCatalog();
            _viewModel.SourceStatus = ConstraintSourceStatusViewModel.From(
                catalogResult,
                !string.IsNullOrWhiteSpace(
                    draft.ConstraintSource.RefDbJsonPath));
            RefreshBindings();
            ActionStatusText.Text = catalogResult.IsUsable
                ? "Quelle erfolgreich geprüft."
                : string.Join(Environment.NewLine, catalogResult.Errors);
        }

        private void SaveClicked(object sender, RoutedEventArgs e)
        {
            CheckSelectionGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            CheckSelectionGrid.CommitEdit(DataGridEditingUnit.Row, true);
            if (_viewModel.PlanCheckSelectionDirty)
            {
                ActionStatusText.Text = "PlanCheck-Auswahl zuerst als versionierten Entwurf übernehmen oder neu laden.";
                return;
            }
            if (_viewModel.Configuration != null && _viewModel.Configuration.HasUnsavedChanges)
            {
                ActionStatusText.Text = "Konfigurationsentwurf zuerst als Version übernehmen oder verwerfen.";
                return;
            }
            AliasesGrid.CommitEdit(DataGridEditingUnit.Cell, true);
            AliasesGrid.CommitEdit(DataGridEditingUnit.Row, true);
            if (_aliasesDirty)
            {
                ActionStatusText.Text = "Aliasentwurf zuerst als Aliasdatei speichern oder mit Neu laden verwerfen.";
                return;
            }
            ClearPlanSettings draft = CreateDraft();
            SettingsValidationResult validation =
                SettingsValidator.Validate(draft, draft.BaseDirectory);
            if (!validation.IsValid)
            {
                ActionStatusText.Text = "Nicht gespeichert:" +
                                        Environment.NewLine +
                                        string.Join(
                                            Environment.NewLine,
                                            validation.Errors);
                return;
            }

            ConstraintCatalogLoadResult catalogResult =
                draft.LoadConstraintCatalog();
            if (!catalogResult.IsUsable)
            {
                ActionStatusText.Text = "Nicht gespeichert:" +
                                        Environment.NewLine +
                                        string.Join(
                                            Environment.NewLine,
                                            catalogResult.Errors);
                return;
            }

            try
            {
                new ConfigurationHistoryStore(draft.ResolvePath(string.IsNullOrWhiteSpace(draft.Paths.ConfigurationDirectory)
                    ? "Configuration" : draft.Paths.ConfigurationDirectory));
                draft.Save();
            }
            catch (Exception exception)
            {
                ActionStatusText.Text = "Einstellungen nicht gespeichert: " + exception.Message;
                return;
            }
            ReloadView();
            ActionStatusText.Text = "Einstellungen gespeichert und Quelle neu geladen.";
            EventHandler handler = SettingsChanged;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        private void ReloadClicked(object sender, RoutedEventArgs e)
        {
            if (!ConfirmDiscardAliases() || !ConfirmDiscardCheckSelection() || !ConfirmDiscardConfiguration()) return;
            ClearPlanSettings.Reload();
            ReloadView();
            ActionStatusText.Text = "Einstellungen von Datenträger neu geladen.";
        }

        private void ReloadView()
        {
            ClearPlanSettings settings = ClearPlanSettings.Reload();
            _viewModel = SettingsViewModel.From(
                settings,
                settings.LoadConstraintCatalog());
            try { _viewModel.InitializeConfiguration(settings); }
            catch (Exception exception)
            {
                ActionStatusText.Text = "Versionsablage nicht zugänglich: " + exception.Message + " Pfad unter Pfade & Quellen korrigieren.";
                SettingsTabs.SelectedIndex = 1;
            }
            var catalog = settings.LoadConstraintCatalog().Catalog;
            _viewModel.SetAliases(catalog == null ? null : catalog.Structures);
            _viewModel.InitializeCheckSelection(_checkSelectionRows);
            _aliasesDirty = false;
            AliasStatusText.Text = ClearPlan.Core.Localization.ReviewLanguage.Text("Angezeigt: aktive Strukturzuordnungen. Änderungen werden nur nach explizitem Speichern aktiviert.");
            DataContext = _viewModel;
        }

        private ClearPlanSettings CreateDraft()
        {
            ClearPlanSettings current = ClearPlanSettings.Load();
            ClearPlanSettings draft =
                JsonConvert.DeserializeObject<ClearPlanSettings>(
                    JsonConvert.SerializeObject(current)) ??
                new ClearPlanSettings();
            _viewModel.ApplyTo(draft);
            return draft;
        }

        private void RefreshBindings()
        {
            DataContext = null;
            DataContext = _viewModel;
        }

        private bool ConfirmDiscardConfiguration()
        {
            return _viewModel == null || _viewModel.Configuration == null || !_viewModel.Configuration.HasUnsavedChanges ||
                MessageBox.Show("Ungespeicherten Konfigurationsentwurf verwerfen?", "Konfigurationen",
                    MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes;
        }

        private void SettingsTabChanged(object sender, SelectionChangedEventArgs e)
        {
            if (e.OriginalSource != SettingsTabs || SettingsTabs.SelectedIndex != 0 || _viewModel == null ||
                _viewModel.Configuration == null || _viewModel.Configuration.HasUnsavedChanges) return;
            _viewModel.Configuration.Refresh();
        }

        private static string BrowseFile(string title, string filter)
        {
            var dialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                CheckFileExists = true,
                Multiselect = false
            };
            return dialog.ShowDialog() == true ? dialog.FileName : null;
        }
    }
}
