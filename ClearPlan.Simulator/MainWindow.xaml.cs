using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.ViewModels;
using Microsoft.Win32;

namespace ClearPlan.Simulator
{
    public partial class MainWindow : Window
    {
        private static readonly string[] CaptureTabIds =
        {
            "overview",
            "pqm",
            "plancheck",
            "fields",
            "dvh",
            "images",
            "parameters",
            "comparison",
            "bev",
            "warnings"
        };

        private readonly SimulatorScenarioRepository repository;
        private readonly SimulatorArguments arguments;
        private readonly SyntheticReportService reportService;
        private readonly VisualCaptureService captureService;
        private readonly SyntheticDvhPngService dvhPngService;
        private readonly ResponsiveLayoutProbeService layoutProbeService;
        private readonly string defaultReportDirectory;
        private bool changingScenario;
        private ReviewSnapshot snapshot;
        private ReviewWorkspaceViewModel workspaceViewModel;

        public MainWindow(
            SimulatorScenarioRepository repository,
            SimulatorArguments arguments,
            SyntheticReportService reportService,
            VisualCaptureService captureService,
            SyntheticDvhPngService dvhPngService,
            ResponsiveLayoutProbeService layoutProbeService)
        {
            this.repository = repository ??
                throw new ArgumentNullException("repository");
            this.arguments = arguments ??
                throw new ArgumentNullException("arguments");
            this.reportService = reportService ??
                throw new ArgumentNullException("reportService");
            this.captureService = captureService ??
                throw new ArgumentNullException("captureService");
            this.dvhPngService = dvhPngService ??
                throw new ArgumentNullException("dvhPngService");
            this.layoutProbeService = layoutProbeService ??
                throw new ArgumentNullException("layoutProbeService");
            defaultReportDirectory =
                SimulatorOutputPaths.GetDefaultReportDirectory(
                    AppDomain.CurrentDomain.BaseDirectory);

            InitializeComponent();
            ConfigureWindowMode();
            ScenarioSelector.ItemsSource =
                new List<string>(repository.ScenarioIds);
            LoadScenario(arguments.ScenarioId);
        }

        private void ConfigureWindowMode()
        {
            if (arguments.IsLayoutProbeMode)
            {
                ApplyLayoutProbeWindowSize();
                return;
            }

            if (!arguments.IsCaptureMode)
            {
                ApplyInteractiveWindowSize();
                return;
            }

            WindowStyle = WindowStyle.None;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = 0;
            Top = 0;
            Width = VisualCaptureService.CaptureWidth;
            Height = VisualCaptureService.CaptureHeight;
            MinWidth = VisualCaptureService.CaptureWidth;
            MinHeight = VisualCaptureService.CaptureHeight;
            MaxWidth = VisualCaptureService.CaptureWidth;
            MaxHeight = VisualCaptureService.CaptureHeight;
            ScenarioSelector.IsEnabled = false;
        }

        private void ApplyLayoutProbeWindowSize()
        {
            WindowStartupLocation = WindowStartupLocation.Manual;
            Left = 0.0;
            Top = 0.0;
            Width = arguments.LayoutProbeWidth;
            Height = arguments.LayoutProbeHeight;
            MinWidth = ReviewWindowSizePolicy.MinimumWidth;
            MinHeight = ReviewWindowSizePolicy.MinimumHeight;
            HorizontalContentAlignment = HorizontalAlignment.Stretch;
            VerticalContentAlignment = VerticalAlignment.Stretch;
            SizeToContent = SizeToContent.Manual;
            ResizeMode = ResizeMode.CanResizeWithGrip;
            ShowInTaskbar = false;
            ScenarioSelector.IsEnabled = false;
        }

        private void ApplyInteractiveWindowSize()
        {
            Rect workArea = SystemParameters.WorkArea;
            ReviewWindowSize windowSize =
                ReviewWindowSizePolicy.Calculate(
                    workArea.Width,
                    workArea.Height);

            Width = windowSize.Width;
            Height = windowSize.Height;
            MinWidth = ReviewWindowSizePolicy.MinimumWidth;
            MinHeight = ReviewWindowSizePolicy.MinimumHeight;
            HorizontalContentAlignment = HorizontalAlignment.Stretch;
            VerticalContentAlignment = VerticalAlignment.Stretch;
            SizeToContent = SizeToContent.Manual;
            ResizeMode = ResizeMode.CanResizeWithGrip;
        }

        private void LoadScenario(string scenarioId)
        {
            ReviewSnapshot nextSnapshot = repository.Load(scenarioId);
            var nextViewModel = new ReviewWorkspaceViewModel(nextSnapshot);
            if (workspaceViewModel != null)
                nextViewModel.Comparison.SetReference(workspaceViewModel.Comparison.ReferenceSnapshot);
            nextViewModel.ReportRequested += OnReportRequested;
            nextViewModel.HtmlReportRequested += OnHtmlReportRequested;
            nextViewModel.OpenPlanRequested += OnOpenPlanRequested;
            nextViewModel.DvhResetRequested += OnDvhResetRequested;
            nextViewModel.DvhExportRequested += OnDvhExportRequested;
            nextViewModel.StructureMappingRequested +=
                OnStructureMappingRequested;

            if (workspaceViewModel != null)
            {
                workspaceViewModel.ReportRequested -= OnReportRequested;
                workspaceViewModel.HtmlReportRequested -= OnHtmlReportRequested;
                workspaceViewModel.OpenPlanRequested -=
                    OnOpenPlanRequested;
                workspaceViewModel.DvhResetRequested -=
                    OnDvhResetRequested;
                workspaceViewModel.DvhExportRequested -=
                    OnDvhExportRequested;
                workspaceViewModel.StructureMappingRequested -=
                    OnStructureMappingRequested;
            }

            snapshot = nextSnapshot;
            workspaceViewModel = nextViewModel;
            ReviewWorkspace.DataContext = workspaceViewModel;
            Title = "ClearPlan Simulator · " + snapshot.ScenarioTitle;
            SimulatorStatusText.Text =
                "Offline · " + snapshot.ScenarioId;

            changingScenario = true;
            ScenarioSelector.SelectedItem = scenarioId;
            changingScenario = false;
        }

        private async void OnLoaded(
            object sender,
            RoutedEventArgs eventArgs)
        {
            try
            {
                VisualCaptureService.SelectTab(
                    ReviewWorkspace,
                    arguments.TabId);
                if (arguments.IsLayoutProbeMode)
                {
                    await layoutProbeService.WriteAsync(
                        this,
                        SimulatorCaptureRoot,
                        ReviewWorkspace,
                        arguments.LayoutProbeWidth,
                        arguments.LayoutProbeHeight,
                        arguments.LayoutProbeOutputPath);
                    Environment.ExitCode = 0;
                    Application.Current.Shutdown(0);
                    return;
                }
                if (arguments.IsCaptureMode)
                {
                    await ExecuteCaptureModeAsync();
                    Environment.ExitCode = 0;
                    Application.Current.Shutdown(0);
                    return;
                }

                if (!string.IsNullOrWhiteSpace(arguments.ReportPath))
                {
                    string report = reportService.Export(
                        snapshot,
                        arguments.ReportPath, ApplyReportOptions);
                    SimulatorStatusText.Text =
                        "Report erstellt · " + Path.GetFileName(report);
                }
            }
            catch (Exception exception)
            {
                App.HandleFatal(exception, arguments.IsAutomatedMode);
            }
        }

        private async Task ExecuteCaptureModeAsync()
        {
            if (!string.IsNullOrWhiteSpace(arguments.CapturePath))
            {
                await captureService.CaptureAsync(
                    SimulatorCaptureRoot,
                    ReviewWorkspace,
                    arguments.TabId,
                    arguments.CapturePath);
                ExportCurrentDvh(
                    GetDeterministicDvhExportPath());
                if (!string.IsNullOrWhiteSpace(arguments.ReportPath))
                {
                    SynchronizeSnapshotDvhSelections();
                    reportService.Export(
                        snapshot,
                        arguments.ReportPath, ApplyReportOptions);
                }
                return;
            }

            string directory = Path.GetFullPath(
                arguments.CaptureAllDirectory);
            SimulatorArguments.ValidateLocalOutputDirectory(
                directory,
                "--capture-all");
            Directory.CreateDirectory(directory);
            foreach (string tabId in CaptureTabIds)
            {
                string outputPath = Path.Combine(
                    directory,
                    tabId + ".png");
                await captureService.CaptureAsync(
                    SimulatorCaptureRoot,
                    ReviewWorkspace,
                    tabId,
                    outputPath);
            }
            ExportCurrentDvh(
                GetDeterministicDvhExportPath());

            string reportPath =
                string.IsNullOrWhiteSpace(arguments.ReportPath)
                    ? Path.Combine(
                        directory,
                        "clearplan-synthetic-report.pdf")
                    : arguments.ReportPath;
            SynchronizeSnapshotDvhSelections();
            reportService.Export(snapshot, reportPath, ApplyReportOptions);
        }

        private void OnScenarioSelectionChanged(
            object sender,
            SelectionChangedEventArgs eventArgs)
        {
            if (changingScenario || !IsLoaded)
            {
                return;
            }

            string scenarioId = ScenarioSelector.SelectedItem as string;
            if (string.IsNullOrWhiteSpace(scenarioId) ||
                string.Equals(
                    scenarioId,
                    snapshot.ScenarioId,
                    StringComparison.Ordinal))
            {
                return;
            }

            try
            {
                LoadScenario(scenarioId);
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    exception.Message,
                    "ClearPlan Simulator",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
                changingScenario = true;
                ScenarioSelector.SelectedItem = snapshot.ScenarioId;
                changingScenario = false;
            }
        }

        private void OnHtmlReportRequested(object sender, ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (!ReferenceEquals(sender, workspaceViewModel) || snapshot == null || !snapshot.Synthetic) return;
            try
            {
                var document = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
                ApplyReportOptions(document);
                foreach (var series in document.DvhSeries)
                {
                    var selection = workspaceViewModel.DvhSeries.FirstOrDefault(row => row.StableId == series.StableId);
                    if (selection != null) series.Selected = selection.IsSelected;
                }
                string html = new ClearPlan.Reporting.MigraDoc.HtmlReviewReportRenderer().Render(document);
                new ClearPlan.Presentation.Views.HtmlReportPreviewWindow(html, defaultReportDirectory) { Owner = this }.Show();
            }
            catch (Exception)
            {
                SimulatorStatusText.Text = "HTML-Quicklook konnte nicht erstellt werden. Die Szenarioansicht bleibt unverändert.";
            }
        }

        private void OnReportRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            Directory.CreateDirectory(defaultReportDirectory);

            var dialog = new SaveFileDialog
            {
                Title = "Synthetischen ClearPlan-Report speichern",
                Filter = "PDF (*.pdf)|*.pdf",
                AddExtension = true,
                DefaultExt = ".pdf",
                InitialDirectory = defaultReportDirectory,
                FileName = snapshot.ScenarioId + "-report.pdf"
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                SynchronizeSnapshotDvhSelections();
                string path = reportService.Export(
                    snapshot,
                    dialog.FileName, ApplyReportOptions);
                SimulatorStatusText.Text =
                    "PDF und HTML-Report erstellt · " + Path.GetFileName(path);
            }
            catch (Exception exception)
            {
                MessageBox.Show(
                    exception.Message,
                    "ClearPlan Simulator",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        private void ApplyReportOptions(ClearPlan.Reporting.ReviewReportDocument document)
        {
            document.IncludeBeamEyeViews = workspaceViewModel.IncludeBeamEyeViews;
            document.HideUnmatched = workspaceViewModel.HideUnmatched;
            document.HiddenStructureIds = workspaceViewModel.HiddenStructureIds.ToList();
        }

        private void OnOpenPlanRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            var plan = eventArgs.Parameter as ReviewPlanRowViewModel;
            if (plan == null)
            {
                SimulatorStatusText.Text =
                    "Plan nicht geöffnet · keine synthetische Auswahl";
                return;
            }

            try
            {
                SimulatorStatusText.Text =
                    SimulatorSnapshotActions.GetOpenPlanStatus(
                        snapshot,
                        plan.PlanKey);
            }
            catch (Exception exception)
            {
                ShowActionError(
                    "Plan konnte nicht geöffnet werden",
                    exception);
            }
        }

        private void OnDvhResetRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            try
            {
                workspaceViewModel.ResetDvhSelections();
                int selected = SynchronizeSnapshotDvhSelections();
                SimulatorStatusText.Text = string.Format(
                    System.Globalization.CultureInfo.InvariantCulture,
                    "DVH zurückgesetzt · {0}/{1} Strukturen ausgewählt",
                    selected,
                    workspaceViewModel.DvhSeries.Count);
            }
            catch (Exception exception)
            {
                ShowActionError(
                    "DVH konnte nicht zurückgesetzt werden",
                    exception);
            }
        }

        private void OnDvhExportRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            if (arguments.IsCaptureMode)
            {
                try
                {
                    ExportCurrentDvh(
                        GetDeterministicDvhExportPath());
                }
                catch (Exception exception)
                {
                    App.HandleFatal(exception, true);
                }
                return;
            }

            Directory.CreateDirectory(defaultReportDirectory);
            var dialog = new SaveFileDialog
            {
                Title = "Synthetisches DVH als PNG speichern",
                Filter = "PNG (*.png)|*.png",
                AddExtension = true,
                DefaultExt = ".png",
                InitialDirectory = defaultReportDirectory,
                FileName = snapshot.ScenarioId + "-dvh.png"
            };
            if (dialog.ShowDialog(this) != true)
            {
                return;
            }

            try
            {
                ExportCurrentDvh(dialog.FileName);
            }
            catch (Exception exception)
            {
                ShowActionError(
                    "DVH konnte nicht exportiert werden",
                    exception);
            }
        }

        private void OnStructureMappingRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            var mapping =
                eventArgs.Parameter as ReviewStructureMappingViewModel;
            if (mapping == null)
            {
                SimulatorStatusText.Text =
                    "Zuordnung nicht übernommen · keine Auswahl";
                return;
            }

            try
            {
                SynchronizeSnapshotDvhSelections();
                string status =
                    SimulatorSnapshotActions.ApplyStructureMapping(
                        snapshot,
                        mapping.StableId,
                        mapping.SelectedStructureId);
                ReviewStructureMapping persisted =
                    snapshot.StructureMappings.Single(
                        row => string.Equals(
                            row.StableId,
                            mapping.StableId,
                            StringComparison.Ordinal));
                workspaceViewModel.MarkStructureMappingApplied(
                    persisted.StableId,
                    persisted.SelectedStructureId,
                    persisted.Message);
                SimulatorStatusText.Text = status;
            }
            catch (Exception exception)
            {
                ShowActionError(
                    "Zuordnung konnte nicht übernommen werden",
                    exception);
            }
        }

        private int SynchronizeSnapshotDvhSelections()
        {
            return SimulatorSnapshotActions.SynchronizeDvhSelections(
                snapshot,
                workspaceViewModel.DvhSeries.Select(
                    row =>
                        new KeyValuePair<string, bool>(
                            row.StableId,
                            row.IsSelected)));
        }

        private string ExportCurrentDvh(string outputPath)
        {
            SynchronizeSnapshotDvhSelections();
            string path = dvhPngService.Export(
                workspaceViewModel.DetailPlotModel,
                outputPath);
            SimulatorStatusText.Text =
                "DVH exportiert · " + Path.GetFileName(path);
            return path;
        }

        private string GetDeterministicDvhExportPath()
        {
            string directory = defaultReportDirectory;
            string fileName = snapshot.ScenarioId + "-dvh.png";
            if (!string.IsNullOrWhiteSpace(
                arguments.CaptureAllDirectory))
            {
                directory = Path.GetFullPath(
                    arguments.CaptureAllDirectory);
                fileName = "clearplan-synthetic-dvh.png";
            }
            else if (!string.IsNullOrWhiteSpace(
                arguments.CapturePath))
            {
                string capturePath = Path.GetFullPath(
                    arguments.CapturePath);
                directory = Path.GetDirectoryName(capturePath);
            }

            if (string.IsNullOrWhiteSpace(directory))
            {
                throw new InvalidOperationException(
                    "The local DVH export directory could not be resolved.");
            }

            SimulatorArguments.ValidateLocalOutputDirectory(
                directory,
                "DVH export directory");
            return string.Equals(
                directory,
                defaultReportDirectory,
                StringComparison.OrdinalIgnoreCase)
                ? SimulatorOutputPaths.GetDefaultDvhExportPath(
                    directory,
                    snapshot.ScenarioId)
                : Path.Combine(directory, fileName);
        }

        private void ShowActionError(
            string actionText,
            Exception exception)
        {
            SimulatorStatusText.Text =
                actionText + " · " + exception.Message;
            MessageBox.Show(
                exception.Message,
                "ClearPlan Simulator",
                MessageBoxButton.OK,
                MessageBoxImage.Error);
        }

        private void OnClosed(
            object sender,
            EventArgs eventArgs)
        {
            if (workspaceViewModel == null)
            {
                return;
            }

            workspaceViewModel.ReportRequested -= OnReportRequested;
            workspaceViewModel.HtmlReportRequested -= OnHtmlReportRequested;
            workspaceViewModel.OpenPlanRequested -= OnOpenPlanRequested;
            workspaceViewModel.DvhResetRequested -= OnDvhResetRequested;
            workspaceViewModel.DvhExportRequested -= OnDvhExportRequested;
            workspaceViewModel.StructureMappingRequested -=
                OnStructureMappingRequested;
        }
    }
}
