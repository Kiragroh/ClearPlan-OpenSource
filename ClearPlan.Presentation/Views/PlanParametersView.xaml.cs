using System.Windows;
using System.Windows.Controls;
namespace ClearPlan.Presentation.Views
{
    public partial class PlanParametersView : UserControl
    {
        public PlanParametersView() { InitializeComponent(); }
        private void ShowParameterPlots(object sender, RoutedEventArgs e)
        {
            var position = ParameterPlotsHeading.TranslatePoint(new Point(0, 0), ParameterScrollViewer);
            ParameterScrollViewer.ScrollToVerticalOffset(ParameterScrollViewer.VerticalOffset + position.Y);
        }
    }
}
