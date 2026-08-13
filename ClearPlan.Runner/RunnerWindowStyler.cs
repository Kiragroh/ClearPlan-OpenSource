using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ClearPlan.Runner
{
    internal static class RunnerWindowStyler
    {
        private const string EssentialsAssemblyName =
            "EsapiEssentials.PluginRunner";

        public static void Register()
        {
            EventManager.RegisterClassHandler(
                typeof(Window),
                FrameworkElement.LoadedEvent,
                new RoutedEventHandler(OnWindowLoaded));
        }

        private static void OnWindowLoaded(object sender, RoutedEventArgs e)
        {
            var window = sender as Window;
            if (window == null || !IsEssentialsWindow(window))
            {
                return;
            }

            try
            {
                window.Title = "ClearPlan Runner";
                window.Background = FindBrush("RunnerCanvasBrush", "#F2F5F9");
                window.Foreground = FindBrush("RunnerTextBrush", "#24324A");
                window.FontFamily = new FontFamily("Segoe UI");
                window.MinWidth = 900;
                window.MinHeight = 600;
                if (double.IsNaN(window.Width) || window.Width < 980)
                {
                    window.Width = 980;
                }
                if (double.IsNaN(window.Height) || window.Height < 700)
                {
                    window.Height = 700;
                }

                window.WindowStartupLocation = WindowStartupLocation.CenterScreen;
                window.Icon = LoadIcon();
            }
            catch
            {
                // Presentation is fail-open; Essentials remains responsible
                // for the clinical runner workflow.
            }
        }

        private static bool IsEssentialsWindow(Window window)
        {
            return string.Equals(
                window.GetType().Assembly.GetName().Name,
                EssentialsAssemblyName,
                StringComparison.Ordinal);
        }

        private static Brush FindBrush(string key, string fallback)
        {
            return Application.Current.TryFindResource(key) as Brush ??
                   (Brush)new BrushConverter().ConvertFromString(fallback);
        }

        private static ImageSource LoadIcon()
        {
            var uri = new Uri(
                "pack://application:,,,/ClearPlan.Runner;component/esapix.ico",
                UriKind.Absolute);
            var info = Application.GetResourceStream(uri);
            if (info == null)
            {
                return null;
            }

            using (info.Stream)
            {
                return BitmapFrame.Create(
                    info.Stream,
                    BitmapCreateOptions.PreservePixelFormat,
                    BitmapCacheOption.OnLoad);
            }
        }
    }
}
