using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ClearPlan.Core.Constraints;
using ClearPlan.Core.Settings;
using ClearPlan.Helpers;
using Microsoft.Win32;
using Newtonsoft.Json;

namespace ClearPlan.Views
{
    public partial class SettingsView : UserControl
    {
        private SettingsViewModel _viewModel;

        public SettingsView()
        {
            InitializeComponent();
            ReloadView();
        }

        public event EventHandler SettingsChanged;

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

            draft.Save();
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
