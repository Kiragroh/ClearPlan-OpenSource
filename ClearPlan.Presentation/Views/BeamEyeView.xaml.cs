using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Presentation.Views
{
    public partial class BeamEyeView : UserControl
    {
        private bool activationQueued;

        public BeamEyeView()
        {
            InitializeComponent();
            DataContextChanged += OnActivationContextChanged;
            IsEnabledChanged += OnActivationContextChanged;
            IsVisibleChanged += OnActivationContextChanged;
        }

        private void OnLoaded(object sender, RoutedEventArgs args) { QueueActivation(); }

        private void OnStatusDetailsClick(object sender, RoutedEventArgs args)
        {
            // A native Button also invokes this for Enter/Space; details are not hover-only.
            BevStatusToolTip.PlacementTarget = BevStatusIndicator;
            BevStatusToolTip.IsOpen = !BevStatusToolTip.IsOpen;
        }

        private void OnActivationContextChanged(object sender, DependencyPropertyChangedEventArgs args) { QueueActivation(); }

        private void QueueActivation()
        {
            if (activationQueued || !IsLoaded || !IsVisible || !IsEnabled) return;
            activationQueued = true;
            // Snapshot replacement raises DataContextChanged before the native host commits its
            // context and leaves its busy guard. Start only after that dispatcher turn finishes.
            Dispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(async () =>
            {
                activationQueued = false;
                if (!IsLoaded || !IsVisible || !IsEnabled) return;
                var model = DataContext as PlanAnalysisViewModel;
                if (model != null) await model.ActivateBevAsync();
            }));
        }
    }
}
