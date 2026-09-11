using System;
using System.Linq;
using ClearPlan.Core.Review;

namespace ClearPlan.Presentation.ViewModels
{
    /// <summary>
    /// Keeps a review workspace on the last validated snapshot and owns the
    /// event subscriptions used by simulator and clinical hosts.
    /// </summary>
    public sealed class ReviewWorkspaceHostController : IDisposable
    {
        private readonly Func<ReviewSnapshot> snapshotFactory;
        private readonly ReviewWorkspaceActionBindings bindings;
        private bool disposed;
        private ReviewSnapshot comparisonReference;

        public ReviewWorkspaceHostController(
            Func<ReviewSnapshot> snapshotFactoryValue,
            ReviewWorkspaceActionBindings bindingsValue,
            ReviewSnapshot comparisonReferenceValue = null)
        {
            if (snapshotFactoryValue == null)
            {
                throw new ArgumentNullException("snapshotFactoryValue");
            }

            snapshotFactory = snapshotFactoryValue;
            bindings = bindingsValue ?? new ReviewWorkspaceActionBindings();
            comparisonReference = comparisonReferenceValue == null ? null :
                ReviewSnapshotJson.Deserialize(ReviewSnapshotJson.Serialize(comparisonReferenceValue));
        }

        public ReviewWorkspaceViewModel CurrentViewModel { get; private set; }

        public ReviewSnapshot CurrentSnapshot { get; private set; }

        public bool TryRefresh(out Exception failure)
        {
            ThrowIfDisposed();
            failure = null;
            try
            {
                return TryActivateSnapshot(
                    snapshotFactory(),
                    out failure);
            }
            catch (Exception exception)
            {
                failure = exception;
                return false;
            }
        }

        public bool TryShowSnapshot(
            ReviewSnapshot snapshot,
            out Exception failure)
        {
            ThrowIfDisposed();
            failure = null;
            try
            {
                return TryActivateSnapshot(snapshot, out failure);
            }
            catch (Exception exception)
            {
                failure = exception;
                return false;
            }
        }

        private bool TryActivateSnapshot(
            ReviewSnapshot snapshot,
            out Exception failure)
        {
            failure = null;
            try
            {
                ReviewSnapshotValidationResult validation =
                    ReviewSnapshotValidator.Validate(snapshot);
                if (!validation.IsValid)
                {
                    string details = string.Join(
                        "; ",
                        validation.Issues
                            .Take(8)
                            .Select(issue =>
                                issue.Code + " at " + issue.Path));
                    throw new InvalidOperationException(
                        "Review snapshot validation failed: " + details);
                }

                var next = new ReviewWorkspaceViewModel(snapshot);
                next.RestorePresentationState(CurrentViewModel);
                if (comparisonReference != null && comparisonReference.Synthetic != snapshot.Synthetic)
                    comparisonReference = null;
                next.Comparison.SetReference(comparisonReference);
                Subscribe(next);

                ReviewWorkspaceViewModel previous = CurrentViewModel;
                CurrentViewModel = next;
                CurrentSnapshot = snapshot;
                if (previous != null)
                {
                    Unsubscribe(previous);
                }

                return true;
            }
            catch (Exception exception)
            {
                failure = exception;
                return false;
            }
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            if (CurrentViewModel != null)
            {
                Unsubscribe(CurrentViewModel);
            }
        }

        private void Subscribe(ReviewWorkspaceViewModel viewModel)
        {
            viewModel.Comparison.ReferenceChanged += OnComparisonReferenceChanged;
            viewModel.ImportPlanDataRequested += bindings.ImportPlanData;
            viewModel.CalculateTargetRequested += bindings.CalculateTarget;
            viewModel.GenerateDrrRequested += bindings.GenerateDrr;
            viewModel.PlanImagesRequested += bindings.PlanImages;
            viewModel.ComparisonPlanRequested += bindings.SelectComparisonPlan;
            viewModel.ReportRequested += bindings.Report;
            viewModel.HtmlReportRequested += bindings.HtmlReport;
            viewModel.AriaUploadRequested += bindings.AriaUpload;
            viewModel.OpenPlanRequested += bindings.OpenPlan;
            viewModel.DvhResetRequested += bindings.ResetDvh;
            viewModel.DvhExportRequested += bindings.ExportDvh;
            viewModel.NavigationRequested += bindings.Navigate;
            viewModel.StructureMappingRequested +=
                bindings.ApplyStructureMapping;
            viewModel.ConstraintTableRequested +=
                bindings.SelectConstraintTable;
            viewModel.SettingsRequested += bindings.OpenSettings;
            viewModel.ScenarioSelectionRequested +=
                bindings.SelectScenario;
        }

        private void Unsubscribe(ReviewWorkspaceViewModel viewModel)
        {
            viewModel.Comparison.ReferenceChanged -= OnComparisonReferenceChanged;
            viewModel.ImportPlanDataRequested -= bindings.ImportPlanData;
            viewModel.CalculateTargetRequested -= bindings.CalculateTarget;
            viewModel.GenerateDrrRequested -= bindings.GenerateDrr;
            viewModel.PlanImagesRequested -= bindings.PlanImages;
            viewModel.ComparisonPlanRequested -= bindings.SelectComparisonPlan;
            viewModel.ReportRequested -= bindings.Report;
            viewModel.HtmlReportRequested -= bindings.HtmlReport;
            viewModel.AriaUploadRequested -= bindings.AriaUpload;
            viewModel.OpenPlanRequested -= bindings.OpenPlan;
            viewModel.DvhResetRequested -= bindings.ResetDvh;
            viewModel.DvhExportRequested -= bindings.ExportDvh;
            viewModel.NavigationRequested -= bindings.Navigate;
            viewModel.StructureMappingRequested -=
                bindings.ApplyStructureMapping;
            viewModel.ConstraintTableRequested -=
                bindings.SelectConstraintTable;
            viewModel.SettingsRequested -= bindings.OpenSettings;
            viewModel.ScenarioSelectionRequested -=
                bindings.SelectScenario;
        }

        private void ThrowIfDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(
                    typeof(ReviewWorkspaceHostController).FullName);
            }
        }

        private void OnComparisonReferenceChanged(object sender, EventArgs eventArgs)
        {
            comparisonReference = ((PlanComparisonViewModel)sender).ReferenceSnapshot;
        }
    }

    public sealed class ReviewWorkspaceActionBindings
    {
        public ReviewWorkspaceActionBindings()
        {
            Report = Ignore;
            HtmlReport = Ignore;
            AriaUpload = Ignore;
            OpenPlan = Ignore;
            ResetDvh = Ignore;
            ExportDvh = Ignore;
            Navigate = Ignore;
            ApplyStructureMapping = Ignore;
            SelectConstraintTable = Ignore;
            OpenSettings = Ignore;
            SelectScenario = Ignore;
            ImportPlanData = Ignore;
            CalculateTarget = Ignore;
            GenerateDrr = Ignore;
            PlanImages = Ignore;
            SelectComparisonPlan = Ignore;
        }

        public EventHandler<ReviewWorkspaceActionEventArgs> Report { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> HtmlReport { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> AriaUpload { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> ImportPlanData { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> CalculateTarget { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> GenerateDrr { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> PlanImages { get; set; }
        public EventHandler<ReviewWorkspaceActionEventArgs> SelectComparisonPlan { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs> OpenPlan { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs> ResetDvh { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs> ExportDvh { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs> Navigate { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs>
            ApplyStructureMapping { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs>
            SelectConstraintTable { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs>
            OpenSettings { get; set; }

        public EventHandler<ReviewWorkspaceActionEventArgs>
            SelectScenario { get; set; }

        private static void Ignore(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
        }
    }

}
