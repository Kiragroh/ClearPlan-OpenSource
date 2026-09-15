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
