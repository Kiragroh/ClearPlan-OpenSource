using System;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewFeedbackPresentationTests
    {
        public static void StatusColorsFollowExplicitSeverity()
        {
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml"));
            string status = Regex.Match(xaml, @"<TextBlock x:Name=""SharedWorkspaceStatusText""(?<body>.*?)</TextBlock>", RegexOptions.Singleline).Value;
            TestAssert.True(status.Contains("ClinicalBlueprintTextMuted"), "Routine status must default to neutral text, not a fixed warning color.");
            foreach (string severity in new[] { "Success", "Warning", "Error" })
                TestAssert.True(status.Contains("Value=\"" + severity + "\""), "Missing explicit status style: " + severity);
            TestAssert.True(status.Contains("ClinicalBlueprintTeal") && status.Contains("ClinicalBlueprintWarningText") && status.Contains("ClinicalBlueprintRed"));
            string code = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            string method = Regex.Match(code, @"internal void ShowAnalysisStatus\((?<body>.*?)\n        \}", RegexOptions.Singleline).Value;
            TestAssert.True(method.Contains("AnalysisStatusSeverity severity = AnalysisStatusSeverity.Info"), "Unclassified progress/info must default to neutral severity.");
            TestAssert.True(method.Contains("SharedWorkspaceStatusText.Tag = severity.ToString()"), "Every status update must replace prior severity.");
            TestAssert.True(method.Contains("details") && method.Contains("Read-only"), "Detailed context and the read-only boundary belong in the tooltip.");
        }

        public static void CompletedAnalysisKeepsLimitationsVisible()
        {
            string code = File.ReadAllText(Path.Combine("ClearPlan.Script", "MainView.xaml.cs"));
            string method = Regex.Match(code, @"internal void ShowCompletedAnalysis\((?<body>.*?)\n        \}", RegexOptions.Singleline).Value;
            TestAssert.True(method.Contains("PamStatus") && method.Contains("GeometryStatus"), "A completion must assess actual PAM and beam geometry availability.");
            TestAssert.True(method.Contains("AnalysisStatusSeverity.Warning") && method.Contains("AnalysisStatusSeverity.Success"));
            TestAssert.True(method.Contains("Geometrie/PAM eingeschränkt"), "Unavailable analysis must remain visible, not only in a tooltip.");
            string host = File.ReadAllText(Path.Combine("ClearPlan.Script", "Review", "ClinicalReviewWorkspaceHost.cs"));
            TestAssert.True(host.Contains("ShowCompletedAnalysis(\"Native ESAPI-Parameter aktualisiert.\", analysis,"), "Native completion needs compact copy plus result-based severity.");
            foreach (Match call in Regex.Matches(host, @"owner.ShowAnalysisStatus\((?<body>.*?)\);", RegexOptions.Singleline))
                TestAssert.True(call.Value.Contains("AnalysisStatusSeverity."), "Host messages need explicit severity: " + call.Value);
        }

        public static void BothDvhTrackersUseOpaqueHighContrastTemplate()
        {
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "ReviewWorkspaceView.xaml"));
            foreach (string name in new[] { "OverviewDvhPlot", "DvhDetailPlot" })
            {
                string plot = Regex.Match(xaml, "<oxy:PlotView x:Name=\"" + name + "\"[^>]*>", RegexOptions.Singleline).Value;
                TestAssert.True(plot.Contains("DefaultTrackerTemplate=\"{StaticResource WorkspaceDvhTrackerTemplate}\""), "Both DVH plots need the same explicit tracker template: " + name);
            }
            string template = Regex.Match(xaml, @"<ControlTemplate x:Key=""WorkspaceDvhTrackerTemplate""(?<body>.*?)</ControlTemplate>", RegexOptions.Singleline).Value;
            TestAssert.True(template.Contains("<oxy:TrackerControl") && template.Contains("Position=\"{Binding Position}\"") && template.Contains("Text=\"{Binding Text}\""));
            TestAssert.True(Regex.IsMatch(template, @"<TextBlock[^>]*Foreground=""\{StaticResource ClinicalBlueprintText\}""", RegexOptions.Singleline), "Tracker text must explicitly override inherited white text.");
            TestAssert.True(template.Contains("Background=\"{StaticResource ClinicalBlueprintSurface}\""), "Tracker background must be an opaque light surface.");
            string theme = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Styles", "ClinicalBlueprint.xaml"));
            double foreground = Luminance(ReadColor(theme, "ClinicalBlueprintText"));
            double background = Luminance(ReadColor(theme, "ClinicalBlueprintSurface"));
            TestAssert.True((Math.Max(foreground, background) + 0.05) / (Math.Min(foreground, background) + 0.05) >= 4.5, "DVH tracker contrast must be at least 4.5:1.");
            double warning = Luminance(ReadColor(theme, "ClinicalBlueprintWarningText"));
            TestAssert.True((background + 0.05) / (warning + 0.05) >= 4.5, "Small status warning text must stay readable on white.");
        }

        private static string ReadColor(string source, string key)
        {
            string value = Regex.Match(source, "x:Key=\"" + key + "\" Color=\"(?<color>#[A-Fa-f0-9]+)\"").Groups["color"].Value;
            TestAssert.Equal(7, value.Length, "Expected an opaque RGB token: " + key);
            return value;
        }

        private static double Luminance(string color)
        {
            var channels = new double[3];
            for (int index = 0; index < 3; index++)
            {
                double channel = int.Parse(color.Substring(1 + index * 2, 2), NumberStyles.HexNumber) / 255.0;
                channels[index] = channel <= 0.04045 ? channel / 12.92 : Math.Pow((channel + 0.055) / 1.055, 2.4);
            }
            return channels[0] * 0.2126 + channels[1] * 0.7152 + channels[2] * 0.0722;
        }
    }
}
