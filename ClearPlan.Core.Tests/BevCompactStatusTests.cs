using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using ClearPlan.Presentation.ViewModels;
using ClearPlan.Rendering;

namespace ClearPlan.Core.Tests
{
    internal static class BevCompactStatusTests
    {
        public static void RunAll()
        {
            CompactRasterOmitsTheMissingProjectionBanner();
            GuiKeepsReasonsAndRecoveryInAccessibleDetails();
            StatusUpdatesWhenTheSelectedProjectionChanges();
        }

        private static void CompactRasterOmitsTheMissingProjectionBanner()
        {
            var render = typeof(BeamEyeViewRenderer).GetMethods().Single(m => m.Name == "Render");
            var compact = render.GetParameters().FirstOrDefault(p => p.Name == "compactStatus");
            TestAssert.NotNull(compact, "The shared renderer needs an explicit compact GUI mode.");
            TestAssert.True(compact.IsOptional && Equals(false, compact.DefaultValue),
                "The PDF/report default must keep its full missing-projection diagnostic.");
            var cp = new ReviewControlPointSample();
            using (var full = Bitmap(BeamEyeViewRenderer.Render(null, cp, null, false)))
            using (var small = Bitmap((byte[])render.Invoke(null, new object[] { null, cp, null, false, 1500, 1000, true })))
            {
                // Sample left of the central ruler: ruler marks must remain visible in both modes.
                TestAssert.True(BrightPixelCount(full, 200, 718, 275, 129) > 1000,
                    "The default report still contains the full missing CT projection notice.");
                TestAssert.Equal(0, BrightPixelCount(small, 200, 718, 275, 129),
                    "The compact viewport must leave that area free of a large missing-CT banner.");
                TestAssert.Equal(0, BrightPixelCount(small, 1030, 710, 440, 55, true),
                    "The GUI rail must not repeat the failure detail below the status symbol.");
            }
            TestAssert.False(BeamEyeViewRenderer.InspectState(null, cp, null, false).GeometryAvailable,
                "Compact presentation must not make unavailable geometry available.");
        }

        private static void GuiKeepsReasonsAndRecoveryInAccessibleDetails()
        {
            var model = Workspace().Analysis;
            model.SelectedControlPoint.GeometryReason = "Missing verified leaf-boundary profile.";
            model.SelectedControlPoint.BevImage = new BeamEyeViewImage
            {
                SourceStatus = "unavailable", UnavailableReason = "CT projection could not be verified."
            };
            model.RenderBevPreview();
            TestAssert.True(model.BevStatusText.Contains("CT projection could not be verified."));
            TestAssert.True(model.BevStatusText.Contains("Missing verified leaf-boundary profile."));
            TestAssert.True(model.BevStatusText.Contains("DRR laden"), "Missing CT details must retain the recovery action.");
            string xaml = Source("ClearPlan.Presentation", "Views", "BeamEyeView.xaml");
            TestAssert.True(xaml.Contains("x:Name=\"BevStatusIndicator\"") && xaml.Contains("Width=\"32\"") && xaml.Contains("Height=\"32\""),
                "The GUI needs a small, keyboard-reachable status control.");
            TestAssert.True(xaml.Contains("AutomationProperties.HelpText=\"{Binding BevStatusText}\"") &&
                xaml.Contains("AutomationProperties.Name=\"{Binding BevStatusSummary}\"") && xaml.Contains("Click=\"OnStatusDetailsClick\""),
                "The status detail must be available to assistive technology and Enter/Space, not hover alone.");
            TestAssert.True(xaml.Contains("<Button.ToolTip>") && xaml.Contains("Content=\"DRR laden\""));
            string tooltip = xaml.Substring(xaml.IndexOf("<Button.ToolTip>", StringComparison.Ordinal));
            tooltip = tooltip.Substring(0, tooltip.IndexOf("</Button.ToolTip>", StringComparison.Ordinal));
            TestAssert.True(tooltip.Contains("Background=\"{StaticResource ClinicalBlueprintSurface}\"") &&
                tooltip.Split(new[] { "Foreground=\"{StaticResource ClinicalBlueprintText}\"" }, StringSplitOptions.None).Length == 4,
                "The tooltip and both text blocks must explicitly use readable theme colors despite host/global styles.");
            TestAssert.False(xaml.Contains("Grid.Row=\"4\" Text=\"{Binding BevStatusText}\""),
                "The lengthy diagnostic must not consume a separate visible footer row.");
            TestAssert.True(Source("ClearPlan.Presentation", "ViewModels", "BeamEyeViewViewModel.cs").Contains("compactStatus: true"),
                "Only the GUI must opt in to the compact renderer.");
        }

        private static void StatusUpdatesWhenTheSelectedProjectionChanges()
        {
            var model = Workspace().Analysis;
            var summary = model.GetType().GetProperty("BevStatusSummary");
            TestAssert.NotNull(summary, "The info control needs a descriptive current status name.");
            model.RenderBevPreview();
            TestAssert.True(((string)summary.GetValue(model, null)).Contains("fehlt"));
            model.SelectedControlPoint.BevImage = new BeamEyeViewImage
            {
                SourceStatus = "available", WidthPixels = 2, HeightPixels = 2, ExtentMm = 100,
                GrayscalePixels = new byte[] { 0, 80, 160, 255 }, ProjectionDescription = "Detached test projection."
            };
            model.RenderBevPreview();
            TestAssert.True(((string)summary.GetValue(model, null)).Contains("verfügbar"));
            TestAssert.True(model.BevStatusText.Contains("Detached test projection."));
            TestAssert.False(model.BevStatusText.Contains("DRR laden"), "Loaded projection must not retain a stale missing-image action.");
        }

        private static ReviewWorkspaceViewModel Workspace()
        {
            var analysis = new ReviewPlanAnalysis();
            var beam = new ReviewBeamAnalysis { BeamId = "Detached status test" };
            beam.ControlPoints.Add(new ReviewControlPointSample());
            analysis.Beams.Add(beam);
            return new ReviewWorkspaceViewModel(new ReviewSnapshot { PlanAnalysis = analysis });
        }

        private static Bitmap Bitmap(byte[] bytes)
        {
            using (var stream = new MemoryStream(bytes))
            using (var decoded = new Bitmap(stream)) return new Bitmap(decoded);
        }

        private static int BrightPixelCount(Bitmap bitmap, int left, int top, int width, int height, bool darkOnWhite = false)
        {
            int count = 0;
            for (int y = top; y < top + height; y++)
                for (int x = left; x < left + width; x++)
                {
                    var pixel = bitmap.GetPixel(x, y);
                    if (darkOnWhite ? pixel.R < 200 : pixel.R > 160 && pixel.G > 160 && pixel.B > 160) count++;
                }
            return count;
        }

        private static string Source(params string[] parts)
        {
            var root = new DirectoryInfo(Directory.GetCurrentDirectory());
            while (root != null && !File.Exists(Path.Combine(root.FullName, "ClearPlan.sln"))) root = root.Parent;
            TestAssert.NotNull(root);
            string path = root.FullName;
            foreach (string part in parts) path = Path.Combine(path, part);
            return File.ReadAllText(path);
        }
    }
}
