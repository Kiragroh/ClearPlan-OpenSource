using System.Windows;
using System.Windows.Controls;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Presentation.Views
{
    public partial class ReviewWorkspaceView : UserControl
    {
        public ReviewWorkspaceView()
        {
            InitializeComponent();
            Loaded += (s,e) => {
                ClearPlan.Core.Localization.ReviewLanguage.Changed += LanguageChanged;
                LanguageChanged(this,System.EventArgs.Empty);
            };
            Unloaded += (s,e) => ClearPlan.Core.Localization.ReviewLanguage.Changed -= LanguageChanged;
        }

        private void LanguageChanged(object sender, System.EventArgs args)
        {
            Dispatcher.Invoke(new System.Action(() => {
                var review=DataContext as ReviewWorkspaceViewModel;
                if (review == null) return;
                ClearPlan.Presentation.Plot.ReviewPlotFactory.RefreshLanguage(review.OverviewPlotModel);
                ClearPlan.Presentation.Plot.ReviewPlotFactory.RefreshLanguage(review.DetailPlotModel);
            }));
        }

        private void OnNavigationSelectionChanged(
            object sender,
            SelectionChangedEventArgs eventArgs)
        {
            if (eventArgs.Source != sender)
            {
                return;
            }

            ReviewWorkspaceViewModel viewModel =
                DataContext as ReviewWorkspaceViewModel;
            TabItem selectedTab =
                ((TabControl)sender).SelectedItem as TabItem;
            if (viewModel == null ||
                selectedTab == null ||
                !viewModel.NavigateCommand.CanExecute(selectedTab.Name))
            {
                return;
            }

            viewModel.NavigateCommand.Execute(selectedTab.Name);
            if (selectedTab.Name == "CollisionTab" && viewModel.IsSynthetic && !viewModel.Collision.HasScene)
                viewModel.Collision.ReloadCommand.Execute(null);
        }
    }
}
