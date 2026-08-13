using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Windows.Input;
using System.Windows.Data;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.Plot;
using OxyPlot;

namespace ClearPlan.Presentation.ViewModels
{
    public sealed class ReviewWorkspaceViewModel : INotifyPropertyChanged
    {
        private string dvhFilterText;
        private bool showSelectedDvhOnly;

        public ReviewWorkspaceViewModel(ReviewSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException("snapshot");
            }

            IsSynthetic = snapshot.Synthetic;
            ModeBadgeText = snapshot.Synthetic
                ? "SYNTHETIC DEMONSTRATION"
                : "CLINICAL READ-ONLY";
            ModeBadgeDescription = snapshot.Synthetic
                ? "NOT FOR CLINICAL USE"
                : "READ-ONLY PLAN REVIEW";
            ScenarioTitle = ReviewDisplayText.ValueOrDash(
                snapshot.ScenarioTitle);
            ScenarioDescription = ReviewDisplayText.ValueOrDash(
                snapshot.ScenarioDescription);
            ProvenanceText = ReviewDisplayText.ValueOrDash(
                snapshot.ProvenanceText);
            PatientDisplayLabel = ReviewDisplayText.ValueOrDash(
                snapshot.PatientDisplayLabel);
            PlanDisplayLabel = ReviewDisplayText.ValueOrDash(
                snapshot.PlanDisplayLabel);
            ActivePlanKey = snapshot.ActivePlanKey;

            Sources = new ObservableCollection<ReviewSourceStatusViewModel>(
                (snapshot.Sources ?? new List<ReviewSourceStatus>())
                    .Select(row => new ReviewSourceStatusViewModel(row)));
            Plans = new ObservableCollection<ReviewPlanRowViewModel>(
                (snapshot.Plans ?? new List<ReviewPlanRow>())
                    .Select(row => new ReviewPlanRowViewModel(row)));
            PqmRows = new ObservableCollection<ReviewPqmRowViewModel>(
                (snapshot.PqmRows ?? new List<ReviewPqmRow>())
                    .Select(row => new ReviewPqmRowViewModel(row)));
            PlanCheckRows =
                new ObservableCollection<ReviewCheckRowViewModel>(
                    (snapshot.PlanCheckRows ?? new List<ReviewCheckRow>())
                        .Select(row => new ReviewCheckRowViewModel(row)));
            FieldRows = new ObservableCollection<ReviewFieldRowViewModel>(
                (snapshot.FieldRows ?? new List<ReviewFieldRow>())
                    .Select(row => new ReviewFieldRowViewModel(row)));
            StructureMappings =
                new ObservableCollection<ReviewStructureMappingViewModel>(
                    (snapshot.StructureMappings ??
                     new List<ReviewStructureMapping>())
                        .Select(
                            row =>
                                new ReviewStructureMappingViewModel(row)));
            DvhSeries =
                new ObservableCollection<ReviewDvhSeriesViewModel>(
                    (snapshot.DvhSeries ?? new List<ReviewDvhSeries>())
                        .Select(row => new ReviewDvhSeriesViewModel(row)));

            dvhFilterText = string.Empty;
            FilteredDvhSeries =
                CollectionViewSource.GetDefaultView(DvhSeries);
            FilteredDvhSeries.Filter = FilterDvhSeries;

            foreach (ReviewDvhSeriesViewModel series in DvhSeries)
            {
                series.PropertyChanged += OnDvhSeriesPropertyChanged;
            }

            OverviewPlotModel = ReviewPlotFactory.Create(DvhSeries);
            DetailPlotModel = ReviewPlotFactory.Create(DvhSeries);

            SourceStatusSummary = BuildSourceStatusSummary(Sources);
            PqmSummary = BuildStatusSummary(PqmRows.Select(row => row.StatusCode));
            PlanCheckSummary = BuildStatusSummary(
                PlanCheckRows.Select(row => row.StatusCode));
            FieldSummary = string.Format(
                CultureInfo.InvariantCulture,
                "{0}/{1} IDs korrekt · {2}/{3} Namen korrekt",
                FieldRows.Count(row => row.IdStatusCode == ReviewStatusCodes.Pass),
                FieldRows.Count,
                FieldRows.Count(
                    row => row.NameStatusCode == ReviewStatusCodes.Pass),
                FieldRows.Count);

            ReportCommand = new RelayCommand(
                parameter => Raise(
                    ReportRequested,
                    ReviewWorkspaceActionCodes.Report,
                    parameter));
            OpenPlanCommand = new RelayCommand(
                parameter => Raise(
                    OpenPlanRequested,
                    ReviewWorkspaceActionCodes.OpenPlan,
                    parameter));
            ResetDvhCommand = new RelayCommand(
                parameter => Raise(
                    DvhResetRequested,
                    ReviewWorkspaceActionCodes.ResetDvh,
                    parameter));
            ExportDvhCommand = new RelayCommand(
                parameter => Raise(
                    DvhExportRequested,
                    ReviewWorkspaceActionCodes.ExportDvh,
                    parameter));
            SelectAllDvhCommand = new RelayCommand(
                parameter => SetAllDvhSelections(true));
            ClearDvhSelectionCommand = new RelayCommand(
                parameter => SetAllDvhSelections(false));
            NavigateCommand = new RelayCommand(
                parameter => Raise(
                    NavigationRequested,
                    ReviewWorkspaceActionCodes.Navigate,
                    parameter));
            ApplyStructureMappingCommand = new RelayCommand(
                parameter => Raise(
                    StructureMappingRequested,
                    ReviewWorkspaceActionCodes.ApplyStructureMapping,
                    parameter));
            SelectConstraintTableCommand = new RelayCommand(
                parameter => Raise(
                    ConstraintTableRequested,
                    ReviewWorkspaceActionCodes.SelectConstraintTable,
                    parameter));
            OpenSettingsCommand = new RelayCommand(
                parameter => Raise(
                    SettingsRequested,
                    ReviewWorkspaceActionCodes.OpenSettings,
                    parameter));
            SelectScenarioCommand = new RelayCommand(
                parameter => Raise(
                    ScenarioSelectionRequested,
                    ReviewWorkspaceActionCodes.SelectScenario,
                    parameter));
        }

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            ReportRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            OpenPlanRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            DvhResetRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            DvhExportRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            NavigationRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            StructureMappingRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            ConstraintTableRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            SettingsRequested;

        public event EventHandler<ReviewWorkspaceActionEventArgs>
            ScenarioSelectionRequested;

        public event PropertyChangedEventHandler PropertyChanged;

        public bool IsSynthetic { get; private set; }

        public string ModeBadgeText { get; private set; }

        public string ModeBadgeDescription { get; private set; }

        public string ScenarioTitle { get; private set; }

        public string ScenarioDescription { get; private set; }

        public string ProvenanceText { get; private set; }

        public string PatientDisplayLabel { get; private set; }

        public string PlanDisplayLabel { get; private set; }

        public string ActivePlanKey { get; private set; }

        public string SourceStatusSummary { get; private set; }

        public string PqmSummary { get; private set; }

        public string PlanCheckSummary { get; private set; }

        public string FieldSummary { get; private set; }

        public ObservableCollection<ReviewSourceStatusViewModel> Sources
        {
            get;
            private set;
        }

        public ObservableCollection<ReviewPlanRowViewModel> Plans
        {
            get;
            private set;
        }

        public ObservableCollection<ReviewPqmRowViewModel> PqmRows
        {
            get;
            private set;
        }

        public ObservableCollection<ReviewCheckRowViewModel> PlanCheckRows
        {
            get;
            private set;
        }

        public ObservableCollection<ReviewFieldRowViewModel> FieldRows
        {
            get;
            private set;
        }

        public ObservableCollection<ReviewStructureMappingViewModel>
            StructureMappings
        {
            get;
            private set;
        }

        public ObservableCollection<ReviewDvhSeriesViewModel> DvhSeries
        {
            get;
            private set;
        }

        public ICollectionView FilteredDvhSeries { get; private set; }

        public string DvhFilterText
        {
            get { return dvhFilterText; }
            set
            {
                string normalizedValue = value ?? string.Empty;
                if (string.Equals(
                    dvhFilterText,
                    normalizedValue,
                    StringComparison.Ordinal))
                {
                    return;
                }

                dvhFilterText = normalizedValue;
                OnPropertyChanged("DvhFilterText");
                FilteredDvhSeries.Refresh();
            }
        }

        public bool ShowSelectedDvhOnly
        {
            get { return showSelectedDvhOnly; }
            set
            {
                if (showSelectedDvhOnly == value)
                {
                    return;
                }

                showSelectedDvhOnly = value;
                OnPropertyChanged("ShowSelectedDvhOnly");
                FilteredDvhSeries.Refresh();
            }
        }

        public PlotModel OverviewPlotModel { get; private set; }

        public PlotModel DetailPlotModel { get; private set; }

        public ICommand ReportCommand { get; private set; }

        public ICommand OpenPlanCommand { get; private set; }

        public ICommand ResetDvhCommand { get; private set; }

        public ICommand ExportDvhCommand { get; private set; }

        public ICommand SelectAllDvhCommand { get; private set; }

        public ICommand ClearDvhSelectionCommand { get; private set; }

        public ICommand NavigateCommand { get; private set; }

        public ICommand ApplyStructureMappingCommand { get; private set; }

        public ICommand SelectConstraintTableCommand { get; private set; }

        public ICommand OpenSettingsCommand { get; private set; }

        public ICommand SelectScenarioCommand { get; private set; }

        public void ResetDvhSelections()
        {
            foreach (ReviewDvhSeriesViewModel series in DvhSeries)
            {
                series.IsSelected = series.IsInitiallySelected;
            }

            FilteredDvhSeries.Refresh();
            OverviewPlotModel.InvalidatePlot(true);
            DetailPlotModel.InvalidatePlot(true);
        }

        public void MarkStructureMappingApplied(
            string stableId,
            string selectedStructureId,
            string message)
        {
            ReviewStructureMappingViewModel mapping =
                StructureMappings.SingleOrDefault(
                    row => string.Equals(
                        row.StableId,
                        stableId,
                        StringComparison.Ordinal));
            if (mapping == null)
            {
                throw new ArgumentException(
                    "The structure mapping is not visible.",
                    "stableId");
            }

            mapping.MarkApplied(
                selectedStructureId,
                message);
        }

        private static string BuildSourceStatusSummary(
            IEnumerable<ReviewSourceStatusViewModel> sources)
        {
            IList<ReviewSourceStatusViewModel> rows = sources.ToList();
            if (rows.Count == 0)
            {
                return "Keine Quellen gemeldet";
            }

            int available = rows.Count(
                row => row.StatusCode == ReviewStatusCodes.Available);
            int fallback = rows.Count(row => row.UsedFallback);
            int unavailable = rows.Count(
                row => row.StatusCode == ReviewStatusCodes.Unavailable);
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} verfügbar · {1} Fallback · {2} nicht verfügbar",
                available,
                fallback,
                unavailable);
        }

        private static string BuildStatusSummary(IEnumerable<string> statuses)
        {
            IList<string> values = statuses.ToList();
            int hintCount = values.Count(
                value => value == ReviewStatusCodes.Variation ||
                         value == ReviewStatusCodes.Info);
            return string.Format(
                CultureInfo.InvariantCulture,
                "{0} erfüllt · {1} {2} · {3} nicht erfüllt",
                values.Count(value => value == ReviewStatusCodes.Pass),
                hintCount,
                hintCount == 1 ? "Hinweis" : "Hinweise",
                values.Count(value => value == ReviewStatusCodes.Fail));
        }

        private void OnDvhSeriesPropertyChanged(
            object sender,
            PropertyChangedEventArgs eventArgs)
        {
            if (eventArgs.PropertyName != "IsSelected")
            {
                return;
            }

            ReviewDvhSeriesViewModel series =
                sender as ReviewDvhSeriesViewModel;
            if (series == null)
            {
                return;
            }

            ReviewPlotFactory.SetSeriesVisibility(
                OverviewPlotModel,
                series.StableId,
                series.IsSelected);
            ReviewPlotFactory.SetSeriesVisibility(
                DetailPlotModel,
                series.StableId,
                series.IsSelected);
            if (ShowSelectedDvhOnly)
            {
                FilteredDvhSeries.Refresh();
            }
        }

        private bool FilterDvhSeries(object item)
        {
            ReviewDvhSeriesViewModel series =
                item as ReviewDvhSeriesViewModel;
            if (series == null)
            {
                return false;
            }

            if (ShowSelectedDvhOnly && !series.IsSelected)
            {
                return false;
            }

            string filter = DvhFilterText.Trim();
            return filter.Length == 0 ||
                   ContainsIgnoreCase(series.DisplayName, filter) ||
                   ContainsIgnoreCase(series.StructureId, filter) ||
                   ContainsIgnoreCase(series.Role, filter);
        }

        private static bool ContainsIgnoreCase(
            string value,
            string searchText)
        {
            return !string.IsNullOrEmpty(value) &&
                   value.IndexOf(
                       searchText,
                       StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private void SetAllDvhSelections(bool isSelected)
        {
            foreach (ReviewDvhSeriesViewModel series in DvhSeries)
            {
                series.IsSelected = isSelected;
            }

            FilteredDvhSeries.Refresh();
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(
                    this,
                    new PropertyChangedEventArgs(propertyName));
            }
        }

        private void Raise(
            EventHandler<ReviewWorkspaceActionEventArgs> handler,
            string actionCode,
            object parameter)
        {
            if (handler != null)
            {
                handler(
                    this,
                    new ReviewWorkspaceActionEventArgs(
                        actionCode,
                        parameter));
            }
        }
    }

    public static class ReviewWorkspaceActionCodes
    {
        public const string Report = "report";
        public const string OpenPlan = "open-plan";
        public const string ResetDvh = "reset-dvh";
        public const string ExportDvh = "export-dvh";
        public const string Navigate = "navigate";
        public const string ApplyStructureMapping = "apply-structure-mapping";
        public const string SelectConstraintTable = "select-constraint-table";
        public const string OpenSettings = "open-settings";
        public const string SelectScenario = "select-scenario";
    }

    public sealed class ReviewWorkspaceActionEventArgs : EventArgs
    {
        public ReviewWorkspaceActionEventArgs(
            string actionCode,
            object parameter)
        {
            ActionCode = actionCode;
            Parameter = parameter;
        }

        public string ActionCode { get; private set; }

        public object Parameter { get; private set; }
    }

    public sealed class ReviewSourceStatusViewModel
    {
        public ReviewSourceStatusViewModel(ReviewSourceStatus row)
        {
            StableId = row.StableId;
            SourceCode = row.SourceCode;
            SourceType = row.SourceType;
            StatusCode = row.Status;
            StatusText = ReviewDisplayText.Status(row.Status);
            OptionalText = row.Optional ? "optional" : "erforderlich";
            UsedFallback = row.UsedFallback;
            PathDisplayLabel = ReviewDisplayText.ValueOrDash(
                row.PathDisplayLabel);
            Message = ReviewDisplayText.ValueOrDash(row.Message);
        }

        public string StableId { get; private set; }

        public string SourceCode { get; private set; }

        public string SourceType { get; private set; }

        public string StatusCode { get; private set; }

        public string StatusText { get; private set; }

        public string OptionalText { get; private set; }

        public bool UsedFallback { get; private set; }

        public string PathDisplayLabel { get; private set; }

        public string Message { get; private set; }
    }

    public sealed class ReviewPlanRowViewModel
    {
        public ReviewPlanRowViewModel(ReviewPlanRow row)
        {
            PlanKey = row.PlanKey;
            DisplayLabel = row.DisplayLabel;
            CreatedText = row.CreatedUtc.HasValue
                ? row.CreatedUtc.Value.ToString(
                    "yyyy-MM-dd HH:mm",
                    CultureInfo.InvariantCulture)
                : "—";
            DosePerFractionText = FormatDose(row.DosePerFractionGy);
            TotalDoseText = FormatDose(row.TotalDoseGy);
            FractionCountText = row.FractionCount.HasValue
                ? row.FractionCount.Value.ToString(
                    CultureInfo.InvariantCulture)
                : "—";
            TargetDisplayLabel = ReviewDisplayText.ValueOrDash(
                row.TargetDisplayLabel);
            StatusCode = row.Status;
            StatusText = ReviewDisplayText.Status(row.Status);
        }

        public string PlanKey { get; private set; }

        public string DisplayLabel { get; private set; }

        public string CreatedText { get; private set; }

        public string DosePerFractionText { get; private set; }

        public string TotalDoseText { get; private set; }

        public string FractionCountText { get; private set; }

        public string TargetDisplayLabel { get; private set; }

        public string StatusCode { get; private set; }

        public string StatusText { get; private set; }

        private static string FormatDose(double? value)
        {
            return value.HasValue
                ? value.Value.ToString(
                    "0.###",
                    CultureInfo.InvariantCulture) + " Gy"
                : "—";
        }
    }

    public sealed class ReviewCheckRowViewModel
    {
        public ReviewCheckRowViewModel(ReviewCheckRow row)
        {
            CheckCode = row.CheckCode;
            Category = row.Category;
            ObservedValue = ReviewDisplayText.ValueOrDash(
                row.ObservedValue);
            ExpectedValue = ReviewDisplayText.ValueOrDash(
                row.ExpectedValue);
            Unit = ReviewDisplayText.ValueOrDash(row.Unit);
            StatusCode = row.Status;
            StatusText = ReviewDisplayText.Status(row.Status);
            SeverityCode = row.Severity;
            SeverityText = ReviewDisplayText.Severity(row.Severity);
            Message = ReviewDisplayText.ValueOrDash(row.Message);
        }

        public string CheckCode { get; private set; }

        public string Category { get; private set; }

        public string ObservedValue { get; private set; }

        public string ExpectedValue { get; private set; }

        public string Unit { get; private set; }

        public string StatusCode { get; private set; }

        public string StatusText { get; private set; }

        public string SeverityCode { get; private set; }

        public string SeverityText { get; private set; }

        public string Message { get; private set; }
    }

    public sealed class ReviewFieldRowViewModel
    {
        public ReviewFieldRowViewModel(ReviewFieldRow row)
        {
            StableId = row.StableId;
            TreatmentOrder = row.TreatmentOrder;
            BeamNumber = row.BeamNumber;
            CurrentId = ReviewDisplayText.ValueOrDash(row.CurrentId);
            ExpectedId = ReviewDisplayText.ValueOrDash(row.ExpectedId);
            CurrentName = ReviewDisplayText.ValueOrDash(row.CurrentName);
            SuggestedName = ReviewDisplayText.ValueOrDash(
                row.SuggestedName);
            IdStatusCode = row.IdStatus;
            IdStatusText = row.IdStatus == ReviewStatusCodes.Pass
                ? "ID okay"
                : "ID prüfen";
            NameStatusCode = row.NameStatus;
            NameStatusText = row.NameStatus == ReviewStatusCodes.Pass
                ? "Name okay"
                : "Name prüfen";
        }

        public string StableId { get; private set; }

        public int TreatmentOrder { get; private set; }

        public int BeamNumber { get; private set; }

        public string CurrentId { get; private set; }

        public string ExpectedId { get; private set; }

        public string CurrentName { get; private set; }

        public string SuggestedName { get; private set; }

        public string IdStatusCode { get; private set; }

        public string IdStatusText { get; private set; }

        public string NameStatusCode { get; private set; }

        public string NameStatusText { get; private set; }
    }

    public sealed class ReviewStructureMappingViewModel :
        INotifyPropertyChanged
    {
        private string selectedStructureId;

        public ReviewStructureMappingViewModel(ReviewStructureMapping row)
        {
            StableId = row.StableId;
            TemplateStructure = row.TemplateStructure;
            selectedStructureId = row.SelectedStructureId;
            AvailableStructureIds = new List<string>(
                row.AvailableStructureIds ?? new List<string>());
            StatusCode = row.Status;
            StatusText = ReviewDisplayText.Status(row.Status);
            Message = ReviewDisplayText.ValueOrDash(row.Message);
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public string StableId { get; private set; }

        public string TemplateStructure { get; private set; }

        public IList<string> AvailableStructureIds { get; private set; }

        public string StatusCode { get; private set; }

        public string StatusText { get; private set; }

        public string Message { get; private set; }

        public string SelectedStructureId
        {
            get { return selectedStructureId; }
            set
            {
                if (selectedStructureId == value)
                {
                    return;
                }

                selectedStructureId = value;
                PropertyChangedEventHandler handler = PropertyChanged;
                if (handler != null)
                {
                    handler(
                        this,
                        new PropertyChangedEventArgs(
                            "SelectedStructureId"));
                }
            }
        }

        public void MarkApplied(
            string selectedStructureIdValue,
            string messageValue)
        {
            SelectedStructureId = selectedStructureIdValue;
            StatusCode = ReviewStatusCodes.Pass;
            StatusText = ReviewDisplayText.Status(
                ReviewStatusCodes.Pass);
            Message = ReviewDisplayText.ValueOrDash(messageValue);
            OnPropertyChanged("StatusCode");
            OnPropertyChanged("StatusText");
            OnPropertyChanged("Message");
        }

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(
                    this,
                    new PropertyChangedEventArgs(propertyName));
            }
        }
    }

    internal static class ReviewDisplayText
    {
        public static string Status(string status)
        {
            switch (status)
            {
                case ReviewStatusCodes.Pass:
                    return "Erfüllt";
                case ReviewStatusCodes.Variation:
                    return "Abweichung";
                case ReviewStatusCodes.Fail:
                    return "Nicht erfüllt";
                case ReviewStatusCodes.Info:
                    return "Hinweis";
                case ReviewStatusCodes.NotEvaluated:
                    return "Nicht bewertet";
                case ReviewStatusCodes.Available:
                    return "Verfügbar";
                case ReviewStatusCodes.Fallback:
                    return "Fallback";
                case ReviewStatusCodes.Unavailable:
                    return "Nicht verfügbar";
                case ReviewStatusCodes.NotConfigured:
                    return "Nicht konfiguriert";
                default:
                    return ValueOrDash(status);
            }
        }

        public static string Severity(string severity)
        {
            switch (severity)
            {
                case ReviewSeverityCodes.None:
                    return "Keine";
                case ReviewSeverityCodes.Info:
                    return "Info";
                case ReviewSeverityCodes.Warning:
                    return "Warnung";
                case ReviewSeverityCodes.Error:
                    return "Fehler";
                default:
                    return ValueOrDash(severity);
            }
        }

        public static string ValueOrDash(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "—" : value;
        }
    }
}
