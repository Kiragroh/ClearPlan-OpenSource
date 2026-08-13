using System;
using ClearPlan.Helpers;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Presentation.Views;

namespace ClearPlan.Review
{
    /// <summary>
    /// Adapts the detached review workspace to the live, read-only Eclipse
    /// view model. All callbacks are executed synchronously on the WPF/ESAPI
    /// thread owned by MainView.
    /// </summary>
    internal sealed class ClinicalReviewWorkspaceHost : IDisposable
    {
        private readonly MainView owner;
        private readonly MainViewModel source;
        private readonly Func<ClearPlanSettings> settingsProvider;
        private readonly ReviewWorkspaceView view;
        private readonly ReviewWorkspaceHostController controller;
        private bool disposed;

        public ClinicalReviewWorkspaceHost(
            MainView ownerValue,
            MainViewModel sourceValue,
            Func<ClearPlanSettings> settingsProviderValue,
            ReviewWorkspaceView viewValue)
        {
            owner = ownerValue ??
                throw new ArgumentNullException("ownerValue");
            source = sourceValue ??
                throw new ArgumentNullException("sourceValue");
            settingsProvider = settingsProviderValue ??
                throw new ArgumentNullException("settingsProviderValue");
            view = viewValue ??
                throw new ArgumentNullException("viewValue");

            var bindings = new ReviewWorkspaceActionBindings
            {
                Report = OnReportRequested,
                OpenPlan = OnOpenPlanRequested,
                ResetDvh = OnDvhResetRequested,
                ExportDvh = OnDvhExportRequested,
                Navigate = OnNavigationRequested,
                ApplyStructureMapping = OnStructureMappingRequested,
                SelectConstraintTable = OnConstraintTableRequested,
                OpenSettings = OnSettingsRequested,
                SelectScenario = OnScenarioSelectionRequested
            };
            controller = new ReviewWorkspaceHostController(
                BuildSnapshot,
                bindings);
        }

        public ReviewWorkspaceViewModel CurrentViewModel
        {
            get { return controller.CurrentViewModel; }
        }

        public bool TryRefresh()
        {
            ThrowIfDisposed();
            Exception failure;
            if (!controller.TryRefresh(out failure))
            {
                owner.ShowSharedWorkspaceFailure(
                    failure,
                    controller.CurrentViewModel != null);
                return false;
            }

            view.DataContext = controller.CurrentViewModel;
            owner.ShowSharedWorkspace();
            return true;
        }

        public void Dispose()
        {
            if (disposed)
            {
                return;
            }

            disposed = true;
            controller.Dispose();
            view.DataContext = null;
        }

        private ClearPlan.Core.Review.ReviewSnapshot BuildSnapshot()
        {
            ClearPlanSettings settings = settingsProvider();
            if (settings == null)
            {
                throw new InvalidOperationException(
                    "ClearPlan settings are unavailable.");
            }

            return new EsapiReviewSnapshotBuilder().Build(
                source,
                settings);
        }

        private void OnReportRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedReport(
                (ReviewWorkspaceViewModel)sender);
        }

        private void OnOpenPlanRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedOpenPlan(
                eventArgs.Parameter as ReviewPlanRowViewModel);
        }

        private void OnDvhResetRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedDvhReset(
                (ReviewWorkspaceViewModel)sender);
        }

        private void OnDvhExportRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedDvhExport(
                (ReviewWorkspaceViewModel)sender);
        }

        private void OnNavigationRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedNavigation(
                eventArgs.Parameter as string);
        }

        private void OnStructureMappingRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedStructureMapping(
                (ReviewWorkspaceViewModel)sender,
                eventArgs.Parameter as ReviewStructureMappingViewModel);
        }

        private void OnConstraintTableRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedConstraintSelection();
        }

        private void OnSettingsRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            owner.HandleSharedSettings();
        }

        private void OnScenarioSelectionRequested(
            object sender,
            ReviewWorkspaceActionEventArgs eventArgs)
        {
            // Clinical mode has no scenario switch. The event remains
            // subscribed so the shared controller has one lifecycle contract.
        }

        private void ThrowIfDisposed()
        {
            if (disposed)
            {
                throw new ObjectDisposedException(
                    typeof(ClinicalReviewWorkspaceHost).FullName);
            }
        }
    }
}
