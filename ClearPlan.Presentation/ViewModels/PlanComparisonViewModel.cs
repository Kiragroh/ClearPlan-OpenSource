using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Input;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.Plot;
using OxyPlot;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed class PlanComparisonViewModel : INotifyPropertyChanged
    {
        private readonly ReviewSnapshot current;
        public PlanComparisonViewModel(ReviewSnapshot snapshot)
        {
            current = snapshot;
            PinReferenceCommand = new RelayCommand(_ => SetReference(current, true));
            ClearReferenceCommand = new RelayCommand(_ => SetReference(null, true), _ => HasReference);
            LoadExampleCommand = new RelayCommand(_ => SetReference(SyntheticScenarioFactory.Create("baseline-pass"), true), _ => current.Synthetic);
            SetReference(null, false);
        }
        public ReviewSnapshot ReferenceSnapshot { get; private set; }
        public bool HasReference { get { return ReferenceSnapshot != null; } }
        public bool IsSynthetic { get { return current.Synthetic; } }
        public bool IsLoading { get; private set; }
        public void SetLoading(bool loading, string message)
        {
            IsLoading = loading;
            if (!string.IsNullOrWhiteSpace(message)) ContextText = message;
            if (PropertyChanged != null) { PropertyChanged(this, new PropertyChangedEventArgs("IsLoading")); PropertyChanged(this, new PropertyChangedEventArgs("ContextText")); }
            CommandManager.InvalidateRequerySuggested();
        }
        public string ReferenceLabel { get { return HasReference ? ReferenceSnapshot.PlanDisplayLabel + " · " + ReferenceSnapshot.ScenarioTitle : "Noch keine Referenz"; } }
        public string CurrentLabel { get { return current.PlanDisplayLabel + " · " + current.ScenarioTitle; } }
        public string ContextText { get; private set; }
        public PlotModel DvhPlotModel { get; private set; }
        public List<ConstraintComparisonRowViewModel> Constraints { get; private set; }
        public List<MetricComparisonRowViewModel> Parameters { get; private set; }
        public ICommand PinReferenceCommand { get; private set; }
        public ICommand ClearReferenceCommand { get; private set; }
        public ICommand LoadExampleCommand { get; private set; }
        public event EventHandler ReferenceChanged;
        public event PropertyChangedEventHandler PropertyChanged;

        public void SetReference(ReviewSnapshot snapshot, bool notify = false)
        {
            if (snapshot != null && snapshot.Synthetic != current.Synthetic)
                throw new ArgumentException("Clinical and synthetic data must remain separate.");
            // Detached, session-only snapshot. No clinical comparison is written to disk or put in the final-plan report.
            ReferenceSnapshot = snapshot == null ? null : ReviewSnapshotJson.Deserialize(ReviewSnapshotJson.Serialize(snapshot));
            Constraints = ReviewComparison.CompareConstraints(ReferenceSnapshot, current)
                .Select(row => new ConstraintComparisonRowViewModel(row)).ToList();
            Parameters = new List<MetricComparisonRowViewModel>();
            if (HasReference)
            {
                var previous = PlanAnalysisViewModel.CreateMetrics(ReferenceSnapshot.PlanAnalysis);
                foreach (var row in PlanAnalysisViewModel.CreateMetrics(current.PlanAnalysis))
                {
                    var before = previous.First(m => m.Key == row.Key);
                    bool compatible = row.Key != "PAM" || string.Equals(ReferenceSnapshot.PlanAnalysis.TargetStructureId,
                        current.PlanAnalysis.TargetStructureId, StringComparison.OrdinalIgnoreCase);
                    if (row.Key == "Small") compatible = ReferenceSnapshot.PlanAnalysis.SmallApertureThresholdCm2 == current.PlanAnalysis.SmallApertureThresholdCm2;
                    Parameters.Add(new MetricComparisonRowViewModel(row.Name, before.Value, row.Value, row.Unit, compatible));
                }
            }
            var curves = new List<ReviewDvhSeriesViewModel>();
            var matchedIds = new HashSet<string>(HasReference ? ReferenceSnapshot.DvhSeries.Select(d => d.StructureId) : Enumerable.Empty<string>(), StringComparer.OrdinalIgnoreCase);
            foreach (var series in current.DvhSeries.Where(d => matchedIds.Contains(d.StructureId)))
            {
                var before = ReferenceSnapshot.DvhSeries.Where(d => string.Equals(d.StructureId, series.StructureId, StringComparison.OrdinalIgnoreCase)).ToList();
                if (before.Count != 1) continue;
                curves.Add(Curve(before[0], "ref-", " · Referenz", series.ColorHex, "dash", series.Selected));
                curves.Add(Curve(series, "current-", " · Aktuell", series.ColorHex, "solid", series.Selected));
            }
            DvhPlotModel = ReviewPlotFactory.Create(curves);
            DvhPlotModel.Subtitle = "Referenz gestrichelt · aktueller Plan durchgezogen · absolute Dosis in Gy";
            DvhPlotModel.SubtitleFontSize = 11;
            var oldPlan = HasReference ? ReferenceSnapshot.Plans.FirstOrDefault(p => p.PlanKey == ReferenceSnapshot.ActivePlanKey) : null;
            var activePlan = current.Plans.FirstOrDefault(p => p.PlanKey == current.ActivePlanKey);
            bool changedPrescription = oldPlan != null && activePlan != null &&
                (oldPlan.TotalDoseGy != activePlan.TotalDoseGy || oldPlan.FractionCount != activePlan.FractionCount);
            ContextText = !HasReference ? "Referenzplan auswählen und laden. Der aktuelle Plan bleibt geöffnet; die Referenz wird nur für diese Sitzung gespeichert." :
                (changedPrescription ? "Verschreibung / Fraktionierung verschieden. Keine automatische Normalisierung. " : "") +
                "Differenz = aktuell − Referenz; keine automatische Aussage über klinische Überlegenheit. " +
                "DVH nur für eindeutig gleiche Struktur-IDs. Der PDF-Report enthält ausschließlich den aktuellen Plan.";
            foreach (string name in new[] { "HasReference", "ReferenceLabel", "Constraints", "Parameters", "DvhPlotModel", "ContextText" })
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(name));
            CommandManager.InvalidateRequerySuggested();
            if (notify && ReferenceChanged != null) ReferenceChanged(this, EventArgs.Empty);
        }

        private static ReviewDvhSeriesViewModel Curve(ReviewDvhSeries row, string prefix, string suffix, string color, string style, bool selected)
        {
            return new ReviewDvhSeriesViewModel(new ReviewDvhSeries { StableId = prefix + row.StableId,
                StructureId = row.StructureId, DisplayName = row.StructureId + suffix, ColorHex = color,
                LineStyle = style, Selected = selected, Role = row.Role, Points = row.Points, VolumeCc = row.VolumeCc });
        }
    }

    public sealed class ConstraintComparisonRowViewModel
    {
        private readonly ReviewComparisonRow row;
        public ConstraintComparisonRowViewModel(ReviewComparisonRow value) { row = value; }
        public string Structure { get { return row.Structure; } }
        public string Objective { get { return row.Objective; } }
        public string Unit { get { return row.Unit; } }
        public string ReferenceValue { get { return PlanAnalysisViewModel.Number(row.Reference == null ? null : row.Reference.AchievedValue); } }
        public string CurrentValue { get { return PlanAnalysisViewModel.Number(row.Current == null ? null : row.Current.AchievedValue); } }
        public string Difference { get { return PlanAnalysisViewModel.Number(row.Difference, "+0.###;-0.###;0"); } }
        public string ReferenceGoal { get { return Goal(row.Reference); } }
        public string CurrentGoal { get { return Goal(row.Current); } }
        public string Note { get { return row.Note; } }
        private static string Goal(ReviewPqmRow value) { return value == null ? "—" : value.Comparator + " " + PlanAnalysisViewModel.Number(value.Goal); }
    }

    public sealed class MetricComparisonRowViewModel
    {
        public MetricComparisonRowViewModel(string label, double? reference, double? current, string unit, bool compatible)
        {
            Label = label; Reference = PlanAnalysisViewModel.Number(reference); Current = PlanAnalysisViewModel.Number(current); Unit = unit;
            Difference = compatible && ReviewComparison.Finite(reference) && ReviewComparison.Finite(current)
                ? PlanAnalysisViewModel.Number(current - reference, "+0.###;-0.###;0") : "—";
            Note = compatible ? "Deskriptiver Vergleich" : "Zielstruktur / Definition verschieden";
        }
        public string Label { get; private set; }
        public string Reference { get; private set; }
        public string Current { get; private set; }
        public string Difference { get; private set; }
        public string Unit { get; private set; }
        public string Note { get; private set; }
    }
}
