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

        public ReviewWorkspaceHostController(
            Func<ReviewSnapshot> snapshotFactoryValue,
            ReviewWorkspaceActionBindings bindingsValue)
        {
            if (snapshotFactoryValue == null)
            {
                throw new ArgumentNullException("snapshotFactoryValue");
            }

            snapshotFactory = snapshotFactoryValue;
            bindings = bindingsValue ?? new ReviewWorkspaceActionBindings();
        }

        public ReviewWorkspaceViewModel CurrentViewModel { get; private set; }

        public bool TryRefresh(out Exception failure)
        {
            ThrowIfDisposed();
            failure = null;
            try
            {
                ReviewSnapshot snapshot = snapshotFactory();
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
                Subscribe(next);

                ReviewWorkspaceViewModel previous = CurrentViewModel;
                CurrentViewModel = next;
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
            viewModel.ReportRequested += bindings.Report;
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
            viewModel.ReportRequested -= bindings.Report;
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
    }

    public sealed class ReviewWorkspaceActionBindings
    {
        public ReviewWorkspaceActionBindings()
        {
            Report = Ignore;
            OpenPlan = Ignore;
            ResetDvh = Ignore;
            ExportDvh = Ignore;
            Navigate = Ignore;
            ApplyStructureMapping = Ignore;
            SelectConstraintTable = Ignore;
            OpenSettings = Ignore;
            SelectScenario = Ignore;
        }

        public EventHandler<ReviewWorkspaceActionEventArgs> Report { get; set; }

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
