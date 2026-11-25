using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows.Input;
using System.Windows.Media;
using ClearPlan.Core.Review;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed partial class PlanImagesViewModel
    {
        private IsodoseDisplayConfiguration isodoseConfiguration, defaultIsodoses;
        private List<ReviewPlanImage> doseDisplayImages = new List<ReviewPlanImage>();
        private bool doseDisplayDirty = true;
        public event EventHandler IsodosesChanged;
        public event EventHandler IsodoseDefaultsRequested;
        public ObservableCollection<IsodoseLevelEditorRow> IsodoseEditorRows { get; private set; }
        public IsodoseDisplayConfiguration IsodoseConfiguration { get { return isodoseConfiguration.Copy(); } }
        public string IsodoseEditorMessage { get; private set; }
        public string IsodoseAvailabilityText { get { return sourceImages.Any(i => i.DosePlane != null && i.DosePlane.IsValid)
            ? "Stufen aus erfassten Dosiswerten · % der Plan-Gesamtverordnung · nur vorhandene Linien werden gezeigt"
            : "Keine editierbare Dosis-Matrix geladen. Vorhandene Linien können umgefärbt werden; für neue Stufen CT laden."; } }
        public IEnumerable<PlanImageLegendItem> IsodoseScaleItems { get { return (LegendItems ?? new List<PlanImageLegendItem>()).Where(i => i.Kind == "isodose").OrderBy(i => i.DoseGy ?? double.MaxValue); } }
        public bool HasIsodoseScale { get { return IsodoseScaleItems.Any(); } }
        public IEnumerable<PlanImageLegendItem> StructureLegendItems { get { return (LegendItems ?? new List<PlanImageLegendItem>()).Where(i => i.Kind == "structure"); } }
        public ICommand ApplyIsodosesCommand { get; private set; }
        public ICommand ResetIsodosesCommand { get; private set; }
        public ICommand EditIsodoseDefaultsCommand { get; private set; }
        public ICommand AddIsodoseCommand { get; private set; }
        public ICommand RemoveIsodoseCommand { get; private set; }
        private IsodoseLevelEditorRow selectedIsodose;
        public IsodoseLevelEditorRow SelectedIsodose { get { return selectedIsodose; } set { selectedIsodose = value; Changed("SelectedIsodose"); CommandManager.InvalidateRequerySuggested(); } }
        private void InitializeIsodoses()
        {
            defaultIsodoses = IsodoseDisplayConfiguration.CreateDefault(); isodoseConfiguration = defaultIsodoses.Copy();
            LoadEditorRows();
            ApplyIsodosesCommand = new RelayCommand(p => ApplyIsodoseDraft(), p => !IsLoading);
            ResetIsodosesCommand = new RelayCommand(p => SetIsodoseConfiguration(defaultIsodoses, "Geladene Defaults angewendet; keine Settings geändert.", false), p => !IsLoading);
            EditIsodoseDefaultsCommand = new RelayCommand(p => { var handler = IsodoseDefaultsRequested; if (handler != null) handler(this, EventArgs.Empty); });
            AddIsodoseCommand = new RelayCommand(p => {
                var occupied = new HashSet<double>();
                foreach (var row in IsodoseEditorRows) { try { occupied.Add(row.ToLevel().Percent); } catch (FormatException) { } }
                double next = occupied.Count == 0 ? 95 : Math.Min(1000, Math.Ceiling(occupied.Max() / 5) * 5 + 5);
                if (occupied.Contains(next)) next = Enumerable.Range(1,1000).First(n => !occupied.Contains(n));
                var added = new IsodoseLevelEditorRow { PercentText = next.ToString(CultureInfo.InvariantCulture), ColorHex = "#FFFFFF", Enabled = true };
                IsodoseEditorRows.Add(added); SelectedIsodose = added;
                IsodoseEditorMessage = "Neue Stufe: Prozentwert und Farbe einstellen, dann Anwenden. Anzeige bis dahin unverändert.";
                Changed("IsodoseEditorMessage"); CommandManager.InvalidateRequerySuggested();
            }, p => !IsLoading && IsodoseEditorRows.Count < 24);
            RemoveIsodoseCommand = new RelayCommand(p => {
                IsodoseEditorRows.Remove(SelectedIsodose); SelectedIsodose = null;
                IsodoseEditorMessage = "Stufe aus dem Entwurf entfernt. Mit Anwenden in GUI und Report übernehmen.";
                Changed("IsodoseEditorMessage");
            }, p => !IsLoading && SelectedIsodose != null && IsodoseEditorRows.Contains(SelectedIsodose));
            IsodoseEditorMessage = "Änderungen gelten nach Anwenden für die Ansicht und den nächsten PDF-/HTML-Report. Defaults bleiben unverändert.";
        }
        private void LoadEditorRows()
        {
            IsodoseEditorRows = new ObservableCollection<IsodoseLevelEditorRow>(isodoseConfiguration.Levels.OrderBy(l => l.Percent)
                .Select(l => new IsodoseLevelEditorRow { PercentText = l.Percent.ToString("0.###", CultureInfo.InvariantCulture), ColorHex = l.ColorHex, Enabled = l.Enabled }));
            SelectedIsodose = null;
            Changed("IsodoseEditorRows");
        }
        public void SetIsodoseConfiguration(IsodoseDisplayConfiguration configuration, string message, bool asDefaults)
        {
            var validated = configuration.Copy();
            if (asDefaults) defaultIsodoses = validated.Copy();
            isodoseConfiguration = validated; doseDisplayDirty = true;
            LoadEditorRows(); RebuildPlanes();
            IsodoseEditorMessage = message ?? "Isodosenanzeige angewendet; Defaults unverändert.";
            Changed("IsodoseEditorMessage"); Changed("IsodoseConfiguration");
            var handler = IsodosesChanged; if (handler != null) handler(this, EventArgs.Empty);
        }
        private void ApplyIsodoseDraft()
        {
            try
            {
                var candidate = new IsodoseDisplayConfiguration { Levels = IsodoseEditorRows.Select(r => r.ToLevel()).ToList() };
                candidate.Validate();
                // Validate before accepting, including retracing; a failed draft never replaces the last applied state.
                var prepared = sourceImages.Select(candidate.Apply).ToList();
                isodoseConfiguration = candidate.Copy(); doseDisplayImages = prepared; doseDisplayDirty = false;
                RebuildPlanes(); IsodoseEditorMessage = "Angewendet für GUI und nächsten PDF-/HTML-Report. Defaults unverändert.";
                Changed("IsodoseConfiguration");
                var handler = IsodosesChanged; if (handler != null) handler(this, EventArgs.Empty);
            }
            catch (Exception ex) when (ex is FormatException || ex is ArgumentException)
            { IsodoseEditorMessage = ex.Message + " Bisherige Anzeige bleibt aktiv."; }
            Changed("IsodoseEditorMessage");
        }
        private void RefreshDoseDisplay()
        {
            if (!doseDisplayDirty) return;
            doseDisplayImages = sourceImages.Select(isodoseConfiguration.Apply).ToList(); doseDisplayDirty = false;
        }
        public List<ReviewPlanImage> ApplyToReport(IEnumerable<ReviewPlanImage> images)
        {
            return (images ?? Enumerable.Empty<ReviewPlanImage>()).Where(i => i != null && i.PlanKey == planKey).Select(isodoseConfiguration.Apply).ToList();
        }
    }
    public sealed class IsodoseLevelEditorRow : INotifyPropertyChanged
    {
        private string colorHex = "#FFFFFF";
        public string PercentText { get; set; } = "100";
        public bool Enabled { get; set; } = true;
        public event PropertyChangedEventHandler PropertyChanged;
        public string ColorHex { get { return colorHex; } set { colorHex = value; var h = PropertyChanged; if (h != null) { h(this, new PropertyChangedEventArgs("ColorHex")); h(this, new PropertyChangedEventArgs("Color")); } } }
        public Brush Color { get { try { var b = (SolidColorBrush)new BrushConverter().ConvertFromString(ColorHex); b.Freeze(); return b; } catch (Exception) { return Brushes.Transparent; } } }
        public IsodoseDisplayLevel ToLevel()
        {
            double percent;
            if (!double.TryParse(PercentText, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.CurrentCulture, out percent) &&
                !double.TryParse(PercentText, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out percent))
                throw new FormatException("Ungültige Isodosenstufe: Prozentwert eingeben.");
            return new IsodoseDisplayLevel { Percent = percent, ColorHex = (ColorHex ?? "").Trim(), Enabled = Enabled };
        }
    }
}
