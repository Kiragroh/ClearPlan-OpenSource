using System;
using System.Windows;
using System.Windows.Input;
using ClearPlan.Core.Localization;
using ClearPlan.Core.Settings;

namespace ClearPlan.Presentation.ViewModels
{
    /// <summary>Explicit save of the current display preferences as the user's own (2026-11-20).
    /// Only display values are written; no review, report or log content.</summary>
    public static class DisplayPreferenceActions
    {
        public static readonly ICommand SaveForUserCommand = new RelayCommand(p => SaveForUser(), p => DisplayPreferences.Store != null);

        public static void SaveForUser()
        {
            var store = DisplayPreferences.Store;
            if (store == null) return;
            try
            {
                store.SaveUser(DisplayPreferences.Current);
                MessageBox.Show(ReviewLanguage.Text("Anzeige-Einstellungen für dich gespeichert. Sie gelten ab jetzt vor dem globalen Standard; löschen unter Einstellungen › Anzeige."),
                    "ClearPlan", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception)
            {
                MessageBox.Show(ReviewLanguage.Text("Anzeige-Einstellungen konnten nicht gespeichert werden. Schreibrecht im ClearPlan-Ordner prüfen."),
                    "ClearPlan", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
