using System;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Navigation;
using Microsoft.Win32;

namespace ClearPlan.Presentation.Views
{
    /// <summary>Displays renderer-owned, script-free HTML in memory. Never launches an OS browser or shell.</summary>
    public sealed class HtmlReportPreviewWindow : Window
    {
        private readonly string html;
        private readonly string initialDirectory;
        private readonly TextBlock status;
        private WebBrowser browser;
        private bool initialNavigation;
        public bool PreviewLoaded { get; private set; }
        public bool PreviewUnavailable { get; private set; }

        public HtmlReportPreviewWindow(string documentHtml, string reportDirectory)
        {
            if (string.IsNullOrWhiteSpace(documentHtml)) throw new ArgumentException("An HTML report is required.", "documentHtml");
            html = documentHtml;
            initialDirectory = reportDirectory;
            Title = "ClearPlan · HTML-Quicklook";
            Width = Math.Min(1180, SystemParameters.WorkArea.Width);
            Height = Math.Min(880, SystemParameters.WorkArea.Height);
            MinWidth = 640;
            MinHeight = 420;
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            Background = Brushes.White;
            FontFamily = new FontFamily("Segoe UI");
            var root = new DockPanel();
            var toolbar = new DockPanel { Margin = new Thickness(16, 12, 16, 12) };
            DockPanel.SetDock(toolbar, Dock.Top);
            var save = new Button { Content = "HTML speichern …", Padding = new Thickness(14, 8, 14, 8),
                Background = new SolidColorBrush(Color.FromRgb(15, 118, 110)), Foreground = Brushes.White,
                BorderThickness = new Thickness(0), FontWeight = FontWeights.SemiBold };
            save.Click += SaveHtml;
            DockPanel.SetDock(save, Dock.Right);
            toolbar.Children.Add(save);
            status = new TextBlock { Text = "Momentaufnahme · Bereits geladene Daten · Keine Neuberechnung",
                Foreground = new SolidColorBrush(Color.FromRgb(32, 43, 56)), TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 16, 0) };
            toolbar.Children.Add(status);
            root.Children.Add(toolbar);
            try
            {
                browser = new WebBrowser();
                browser.Navigating += BlockNavigation;
                browser.LoadCompleted += (sender, args) => { PreviewLoaded = true; initialNavigation = false; };
                root.Children.Add(browser);
                Loaded += (sender, args) => {
                    try { initialNavigation = true; browser.NavigateToString(html); }
                    catch (Exception) { MarkPreviewUnavailable(); }
                };
            }
            catch (Exception) { MarkPreviewUnavailable(); }
            Content = root;
            Closed += (sender, args) => { if (browser != null) browser.Dispose(); };
        }

        private void BlockNavigation(object sender, NavigatingCancelEventArgs e)
        {
            // NavigateToString uses about:blank. Reject links, files, remote requests,
            // custom protocols and subsequent top-level navigation from the document.
            if (initialNavigation && (e.Uri == null || e.Uri.AbsoluteUri == "about:blank"))
            {
                initialNavigation = false;
                return;
            }
            if (PreviewLoaded && e.Uri != null && e.Uri.AbsoluteUri.StartsWith("about:blank#", StringComparison.Ordinal))
                return; // Same-document section links only; no origin/file/protocol transition.
            e.Cancel = true;
        }

        private void MarkPreviewUnavailable()
        {
            PreviewUnavailable = true;
            status.Text = "HTML-Vorschau auf diesem Arbeitsplatz nicht verfügbar. Der Report kann weiterhin als HTML gespeichert werden.";
        }

        private void SaveHtml(object sender, RoutedEventArgs args)
        {
            var dialog = new SaveFileDialog { Title = "ClearPlan · HTML-Report speichern", Filter = "HTML (*.html)|*.html",
                DefaultExt = ".html", AddExtension = true, FileName = "ClearPlan-plan-review.html",
                InitialDirectory = initialDirectory ?? string.Empty };
            if (dialog.ShowDialog(this) != true) return;
            try
            {
                File.WriteAllText(dialog.FileName, html, new UTF8Encoding(false));
                status.Text = "HTML gespeichert · " + Path.GetFileName(dialog.FileName);
                status.ToolTip = dialog.FileName;
            }
            catch (Exception)
            {
                status.Text = "HTML konnte nicht gespeichert werden. Bitte Speicherort und Zugriffsrechte prüfen.";
            }
        }
    }
}
