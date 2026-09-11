using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;

namespace ClearPlan.Views
{
    public partial class ConfigurationWorkspaceView : UserControl
    {
        private bool _changingSelection;
        private ConfigurationWorkspaceViewModel Model { get { return DataContext as ConfigurationWorkspaceViewModel; } }
        public ConfigurationWorkspaceView() { InitializeComponent(); }

        private bool ConfirmDiscard()
        {
            return Model == null || !Model.HasUnsavedChanges || MessageBox.Show(
                "Ungespeicherten Konfigurationsentwurf verwerfen? Der gespeicherte Dateistand bleibt erhalten.",
                "Konfigurationsentwurf", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes;
        }

        private void EntryChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_changingSelection || Model == null || EntriesList.SelectedItem == Model.SelectedEntry) return;
            _changingSelection = true;
            try
            {
                if (ConfirmDiscard()) Model.Select(EntriesList.SelectedItem as ConfigurationEntryViewModel);
                else EntriesList.SelectedItem = Model.SelectedEntry;
            }
            finally { _changingSelection = false; }
        }

        private void ImportConfiguredClicked(object sender, RoutedEventArgs e)
        {
            if (Model != null && ConfirmDiscard()) Model.ImportFile(Model.ConfiguredPath);
        }

        private void ImportOtherClicked(object sender, RoutedEventArgs e)
        {
            if (Model == null || Model.SelectedEntry == null || !ConfirmDiscard()) return;
            var dialog = new OpenFileDialog { Title = "Als versionierte Arbeitskopie importieren", CheckFileExists = true,
                Multiselect = false, Filter = Model.IsExcel ? "ClearPlan-Constraint-Katalog (*.xlsx)|*.xlsx" : "JSON-Konfiguration (*.json)|*.json" };
            if (dialog.ShowDialog() == true) Model.ImportFile(dialog.FileName);
        }

        private void RefreshClicked(object sender, RoutedEventArgs e) { if (Model != null && ConfirmDiscard()) Model.CancelEdit(); }
        private void CancelClicked(object sender, RoutedEventArgs e) { if (Model != null && ConfirmDiscard()) Model.CancelEdit(); }
        private void SaveJsonClicked(object sender, RoutedEventArgs e) { if (Model != null) Model.SaveJson(); }
        private void AcceptExcelClicked(object sender, RoutedEventArgs e) { if (Model != null) Model.AcceptExcelDraft(); }

        private void OpenExcelClicked(object sender, RoutedEventArgs e)
        {
            if (Model == null) return;
            string path = Model.PrepareExcelDraft();
            if (path == null) return;
            try { Process.Start(new ProcessStartInfo { FileName = path, UseShellExecute = true }); }
            catch (Exception)
            {
                MessageBox.Show("Der Excel-Entwurf konnte nicht geöffnet werden. Excel-Dateizuordnung prüfen. Der aktive Dateistand wurde nicht verändert.",
                    "Excel-Entwurf", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        private void RestoreClicked(object sender, RoutedEventArgs e)
        {
            if (Model == null || Model.SelectedVersion == null || !ConfirmDiscard()) return;
            if (MessageBox.Show("Version " + Model.SelectedVersion.RevisionNumber +
                " als neuen Dateistand wiederherstellen? Die bisherige Historie bleibt erhalten. Aktivierung erst durch Einstellungen speichern und erneutes Laden der Planansicht.",
                "Version wiederherstellen", MessageBoxButton.YesNo, MessageBoxImage.Question, MessageBoxResult.No) == MessageBoxResult.Yes)
                Model.RestoreSelected();
        }
    }
}
