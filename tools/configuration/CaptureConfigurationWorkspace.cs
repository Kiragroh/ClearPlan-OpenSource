using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using ClearPlan;
using ClearPlan.Core.Configuration;
using ClearPlan.Views;

// Offscreen, synthetic-only native WPF rendering; no TPS startup and no window activation.
internal static class CaptureConfigurationWorkspace
{
    [STAThread]
    public static int Main(string[] args)
    {
        if (args.Length != 2) throw new ArgumentException("Usage: capture <repository-root> <output-directory>");
        string repo = Path.GetFullPath(args[0]);
        string output = Path.GetFullPath(args[1]);
        Directory.CreateDirectory(output);
        string configurationRoot = Path.Combine(output, "configuration-capture-" + Guid.NewGuid().ToString("N"));
        string source = Path.Combine(repo, "ClearPlan.Script", "Distribution", "FieldNamingRules.json");
        var store = new ConfigurationHistoryStore(configurationRoot);
        byte[] data = File.ReadAllBytes(source);
        store.Import("field-naming", "Default · Feldnamen", ".json", data, "SYNTHETIC\\reviewer", "Synthetischer Ausgangsstand");
        store.Save("field-naming", data, "SYNTHETIC\\reviewer", "Feldnamen-Nomenklatur geprüft", 1);
        var entries = new List<ConfigurationEntryViewModel>
        {
            Entry("default-rules", "Default-Zielregeln", ".json", source),
            Entry("constraints", "Constraints · Excel", ".xlsx", source),
            Entry("aliases", "Strukturnamen & Aliase", ".json", source),
            Entry("field-naming", "Default · Feldnamen", ".json", source),
            Entry("mlc-profiles", "MLC-Geometrieprofile", ".json", source),
            Entry("refdb", "RefDB · optionale Kopie", ".json", source)
        };
        var model = new ConfigurationWorkspaceViewModel(configurationRoot, entries, (key, path) => { });
        model.Select(entries[3]);
        var app = new Application();
        var view = new ConfigurationWorkspaceView { DataContext = model };
        Capture(view, 1280, 900, Path.Combine(output, "configuration-1280.png"));
        Capture(view, 1600, 1000, Path.Combine(output, "configuration-1600.png"));
        var scroll = Descendants(view).OfType<ScrollViewer>()
            .Where(item => item.ScrollableHeight > 0).OrderByDescending(item => item.ActualHeight).First();
        scroll.ScrollToBottom();
        Capture(view, 1600, 1000, Path.Combine(output, "configuration-history-1600.png"));
        return 0;
    }

    private static ConfigurationEntryViewModel Entry(string key, string title, string extension, string path)
    {
        return new ConfigurationEntryViewModel { Key = key, Title = title, Extension = extension,
            Description = "Synthetische Darstellung · keine Patientendaten. Beispiel: 120-30 T300 GUZ – Tisch vorletzter Bestandteil, Richtung letzter. Nur Namensvorschau, keine Feldänderung.",
            SourcePath = () => path, Validate = bytes => { } };
    }

    private static void Capture(FrameworkElement view, int width, int height, string path)
    {
        view.Width = width; view.Height = height;
        view.Measure(new Size(width, height));
        view.Arrange(new Rect(0, 0, width, height));
        view.UpdateLayout();
        var bitmap = new RenderTargetBitmap(width, height, 96, 96, PixelFormats.Pbgra32);
        bitmap.Render(view);
        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(bitmap));
        using (var stream = new FileStream(path, FileMode.CreateNew)) encoder.Save(stream);
    }

    private static IEnumerable<DependencyObject> Descendants(DependencyObject parent)
    {
        for (int index = 0; index < VisualTreeHelper.GetChildrenCount(parent); index++)
        {
            var child = VisualTreeHelper.GetChild(parent, index);
            yield return child;
            foreach (var descendant in Descendants(child)) yield return descendant;
        }
    }
}
