using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using OxyPlot;
using OxyPlot.Axes;
using OxyPlot.Series;
using ClearPlan.Core.Localization;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed partial class PlanAnalysisViewModel : INotifyPropertyChanged
    {
        private ReviewBeamAnalysis selectedBeam;
        private string selectedTargetStructureId;
        public PlanAnalysisViewModel(ReviewPlanAnalysis analysis, bool synthetic = false)
            : this(analysis, synthetic, null) { }

        public PlanAnalysisViewModel(ReviewPlanAnalysis analysis, bool synthetic,
            Func<ReviewBeamAnalysis, ReviewControlPointSample, BeamEyeViewImage> syntheticDrrProvider)
        {
            this.syntheticDrrProvider = syntheticDrrProvider ?? ClearPlan.Core.Simulation.SyntheticDrrFactory.Create;
            Analysis = analysis ?? new ReviewPlanAnalysis();
            Beams = Analysis.Beams ?? new List<ReviewBeamAnalysis>();
            Metrics = CreateMetrics(Analysis);
            AvailableTargets = Analysis.AvailableTargetStructureIds ?? new List<string>();
            selectedTargetStructureId = Analysis.TargetStructureId;
            SelectedBeam = Beams.FirstOrDefault();
            if (selectedBeam == null) RefreshPlots();
            InitializeBev(synthetic);
        }

        public ReviewPlanAnalysis Analysis { get; private set; }
        public List<ReviewBeamAnalysis> Beams { get; private set; }
        public List<AnalysisMetricRow> Metrics { get; private set; }
        public List<ReviewTargetQuality> TargetQuality { get { return Analysis.TargetQuality ?? new List<ReviewTargetQuality>(); } }
        public string TargetQualityNote { get { return Analysis.TargetQualityNote ?? "PTV-Qualitätsmetriken noch nicht verfügbar. Native Plandaten aktualisieren; keine TPS-Änderungen."; } }
        public List<string> AvailableTargets { get; private set; }
        public System.Windows.Input.ICommand ImportPlanDataCommand { get; internal set; }
        public System.Windows.Input.ICommand CalculateTargetCommand { get; internal set; }
        public string SelectedTargetStructureId
        {
            get { return selectedTargetStructureId; }
            set { selectedTargetStructureId = value; Changed("SelectedTargetStructureId"); }
        }
        public string TotalMuText { get { return Number(Analysis.TotalMetersetMu, "0.0") + " MU"; } }
        public string MuPerGyText { get { return Number(Analysis.MuPerGy, "0.0") + " MU/Gy"; } }
        public string PamText { get { return Number(Analysis.Pam, "0.000"); } }
        public string AreaText { get { return Number(Analysis.MeanApertureAreaCm2, "0.00") + " cm²"; } }
        public string TargetText { get { return string.IsNullOrWhiteSpace(Analysis.TargetStructureId) ? "Keine eindeutige Zielstruktur" : Analysis.TargetStructureId; } }
        public string MethodText { get { return "PAM: verdeckter Anteil der Zielprojektion (0–1), innerhalb eines Feldes mit Kontrollpunkt-Meterset gewichtet. " +
            (Analysis.PamWeightingMode == "BeamWeightFactor" ? "Feldaggregation wie PlanCheck mit nativen Beam.WeightFactor-Werten. " : "Feldaggregation mit MU. ") +
            "Höher bedeutet stärkere Aperturmodulation, nicht automatisch schlechtere Planqualität.\n\n" +
            (Analysis.TargetSelectionProvenance ?? "") + "\n" + (Analysis.PamReason ?? "") + "\n" + (Analysis.GeometryProvenance ?? ""); } }
        public string WarningText { get { return string.Join("\n", Analysis.Warnings ?? new List<string>()); } }
        public string DoseRateText { get { return (selectedBeam == null ? "" : selectedBeam.DoseRateEstimateReason + "\n") +
            (Analysis.DoseRateProvenance ?? "Keine gemessenen Lieferdaten. Fehlende Werte werden nicht als Null dargestellt."); } }
        public string DoseRateSummary { get { return HasEstimatedDoseRate
            ? "Geschätzter Planverlauf · keine gemessene Abgabe"
            : HasPlannedDoseRate ? "Gelieferte Planwerte · keine gemessene Abgabe" : "Kein auswertbarer Dosisratenverlauf"; } }
        public string DoseRateProfileText { get { return HasEstimatedDoseRate
            ? "Profil: " + selectedBeam.DoseRateEstimateProfile + " · Gantry-Annahme " +
                Number(selectedBeam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond, "0.##") + " °/s"
            : "Schätzprofil unter Einstellungen prüfen; Details im Tooltip."; } }
        public string NominalDoseRateText { get { return "Nominale Einstellung: " +
            Number(selectedBeam == null ? null : selectedBeam.NominalDoseRateMuPerMin, "0.##") + " MU/min"; } }
        public string BeamText { get { return selectedBeam == null ? "Kein auswertbares Behandlungsfeld" : selectedBeam.MachineId + " · " + selectedBeam.MlcModel + " · " + selectedBeam.MlcLayerCount + " MLC-Lage(n) · " + (selectedBeam.HasJaws == true ? "mit Blenden" : selectedBeam.HasJaws == false ? "ohne Blenden" : "Blendenstatus unbekannt"); } }
        public ReviewBeamAnalysis SelectedBeam
        {
            get { return selectedBeam; }
            set
            {
                selectedBeam = value;
                RefreshPlots();
                Changed("SelectedBeam"); Changed("BeamText"); Changed("NominalDoseRateText"); Changed("DoseRateText");
                ResetBevSelection();
            }
        }
        public PlotModel DoseRatePlotModel { get; private set; }
        public bool HasPlannedDoseRate { get { return DoseRatePlotModel != null && DoseRatePlotModel.Series.Any(s => (string)s.Tag == "planned"); } }
        public bool HasEstimatedDoseRate { get { return DoseRatePlotModel != null && DoseRatePlotModel.Series.Any(s => (string)s.Tag == "estimated"); } }
        public bool HasDoseRateTrace { get { return HasPlannedDoseRate || HasEstimatedDoseRate; } }
        public PlotModel AperturePlotModel { get; private set; }
        public event PropertyChangedEventHandler PropertyChanged;
        private void Changed(string name) { if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs(name)); }

        public void RefreshLanguage() { RefreshPlots(); }

        private void RefreshPlots()
        {
            DoseRatePlotModel = CreatePlot("Kontrollpunkt", "Dosisrate [MU/min]");
            AperturePlotModel = CreatePlot("Kontrollpunkt", "Aperturfläche [cm²]");
            var planned = Line("Geplant (Kontrollpunkt)", "#0F766E", LineStyle.Solid);
            planned.Tag = "planned";
            var estimated = Line("Geschätzt (Segment)", "#0F766E", LineStyle.Dash);
            estimated.Tag = "estimated";
            var area = Line("Effektive Apertur", "#0F766E", LineStyle.Solid);
            foreach (var point in selectedBeam == null || selectedBeam.ControlPoints == null ? new List<ReviewControlPointSample>() : selectedBeam.ControlPoints)
            {
                if (point == null) continue;
                // Beam.DoseRate is a scalar setting, not a rate trajectory. Without a
                // time basis, meterset weights cannot provide MU/min for RapidArc.
                AddPoint(planned, point.Index, point.PlannedDoseRateMuPerMin);
                AddPoint(estimated, point.Index, selectedBeam.DoseRateEstimateStatus == "Estimated" ? point.EstimatedDoseRateMuPerMin : null);
                AddPoint(area, point.Index, point.ApertureAreaCm2);
            }
            if (planned.Points.Any(p => !double.IsNaN(p.Y))) DoseRatePlotModel.Series.Add(planned);
            if (estimated.Points.Any(p => !double.IsNaN(p.Y))) DoseRatePlotModel.Series.Add(estimated);
            if (area.Points.Any(p => !double.IsNaN(p.Y))) AperturePlotModel.Series.Add(area);
            DoseRatePlotModel.Subtitle = ReviewLanguage.Text(HasDoseRateTrace ? "MU/min · keine gemessene Abgaberate" : "Kein Dosisratenverlauf verfügbar");
            AperturePlotModel.Subtitle = ReviewLanguage.Text(AperturePlotModel.Series.Count == 0 ? "Vollständige Aperturgeometrie nicht verfügbar" : "Schnittmenge aller MLC-Lagen und vorhandener Blenden");
            Changed("DoseRatePlotModel"); Changed("AperturePlotModel");
            Changed("HasPlannedDoseRate");
            Changed("HasEstimatedDoseRate"); Changed("HasDoseRateTrace"); Changed("DoseRateSummary"); Changed("DoseRateProfileText");
        }

        private static void AddPoint(LineSeries line, int index, double? value)
        {
            line.Points.Add(new DataPoint(index, ReviewComparison.Finite(value) && value.Value >= 0 ? value.Value : double.NaN));
        }
        private static LineSeries Line(string title, string color, LineStyle style)
        { return new LineSeries { Title = ReviewLanguage.Text(title), Color = OxyColor.Parse(color), LineStyle = style, StrokeThickness = 2.2,
            MarkerType = MarkerType.Circle, MarkerSize = 2.5, MarkerFill = OxyColor.Parse(color), MarkerStroke = OxyColor.Parse(color) }; }

        internal static PlotModel CreatePlot(string xLabel, string yLabel)
        {
            var plot = new PlotModel
            {
                Background = OxyColors.White, TextColor = OxyColor.Parse("#344454"), DefaultFont = "Segoe UI",
                DefaultFontSize = 12, SubtitleFontSize = 11, SubtitleColor = OxyColor.Parse("#586777"),
                PlotAreaBorderColor = OxyColor.Parse("#D6DDE4"), LegendPosition = LegendPosition.TopRight,
                LegendPlacement = LegendPlacement.Outside, LegendOrientation = LegendOrientation.Horizontal,
                LegendFontSize = 11, LegendBorderThickness = 0
            };
            plot.Axes.Add(new LinearAxis { Position = AxisPosition.Bottom, Title = ReviewLanguage.Text(xLabel), Minimum = 0,
                IsZoomEnabled = false, IsPanEnabled = false,
                MajorGridlineStyle = LineStyle.None, AxislineColor = OxyColor.Parse("#BAC5CE") });
            plot.Axes.Add(new LinearAxis { Position = AxisPosition.Left, Title = ReviewLanguage.Text(yLabel), Minimum = 0,
                IsZoomEnabled = false, IsPanEnabled = false,
                MajorGridlineStyle = LineStyle.Solid, MajorGridlineColor = OxyColor.Parse("#E8EDF0"),
                AxislineColor = OxyColor.Parse("#BAC5CE") });
            return plot;
        }

        public static string Number(double? value, string format = "0.###")
        { return ReviewComparison.Finite(value) ? value.Value.ToString(format, CultureInfo.InvariantCulture) : "—"; }

        public static List<AnalysisMetricRow> CreateMetrics(ReviewPlanAnalysis value)
        {
            return new List<AnalysisMetricRow>
            {
                new AnalysisMetricRow("MU", "Gesamt-MU", value.TotalMetersetMu, "MU", "Nur Behandlungsfelder; keine Setup-Felder"),
                new AnalysisMetricRow("MU/Gy", "MU pro Fraktionsdosis", value.MuPerGy, "MU/Gy", "Gesamt-MU / verschriebene Dosis pro Fraktion"),
                new AnalysisMetricRow("Normalization", "Plannormierung", value.PlanNormalizationPercent, "%", "TPS-Einstellung; beschreibt keine klinische Überlegenheit und ersetzt keinen Dosisvergleich"),
                new AnalysisMetricRow("PAM", "Plan Aperture Modulation", value.Pam, "1", value.PamReason ?? "Zielbezogen; MU-gewichteter verdeckter Projektionsanteil"),
                new AnalysisMetricRow("Area", "Mittlere Aperturfläche", value.MeanApertureAreaCm2, "cm²", "MU-gewichtet; effektive Öffnung aller Lagen"),
                new AnalysisMetricRow("Small", "Anteil kleiner Öffnungen", value.SmallApertureFraction, "1", "MU-gewichteter CP-Anteil mit Fläche < " + Number(value.SmallApertureThresholdCm2) + " cm²; kein SAS10")
            };
        }
    }

    public sealed class AnalysisMetricRow
    {
        public AnalysisMetricRow(string key, string name, double? value, string unit, string note)
        { Key = key; Name = name; Value = value; Unit = unit; Note = note; }
        public string Key { get; private set; }
        public string Name { get; private set; }
        public double? Value { get; private set; }
        public string ValueText { get { return PlanAnalysisViewModel.Number(Value); } }
        public string Unit { get; private set; }
        public string Note { get; private set; }
    }
}
