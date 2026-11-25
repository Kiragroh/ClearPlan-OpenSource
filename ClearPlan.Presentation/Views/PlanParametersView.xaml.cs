using System.Windows;
using System.Windows.Controls;
namespace ClearPlan.Presentation.Views
{
    public partial class PlanParametersView : UserControl
    {
        public PlanParametersView()
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
                var model=DataContext as ClearPlan.Presentation.ViewModels.PlanAnalysisViewModel;
                if(model != null) model.RefreshLanguage();
            }));
        }
        private void ShowParameterPlots(object sender, RoutedEventArgs e)
        {
            var position = ParameterPlotsHeading.TranslatePoint(new Point(0, 0), ParameterScrollViewer);
            ParameterScrollViewer.ScrollToVerticalOffset(ParameterScrollViewer.VerticalOffset + position.Y);
        }
    }
}
