using System;
using System.IO;
using System.Linq;
using OxyPlot;
using OxyPlot.Wpf;

namespace ClearPlan.Simulator
{
    public sealed class SyntheticDvhPngService
    {
        public const int ExportWidth = 1400;
        public const int ExportHeight = 800;

        public string Export(
            PlotModel plotModel,
            string outputPath)
        {
            if (plotModel == null)
            {
                throw new ArgumentNullException("plotModel");
            }

            SimulatorArguments.ValidateLocalOutputFile(
                outputPath,
                ".png",
                "DVH PNG");
            string fullPath = Path.GetFullPath(outputPath);
            string parent = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(parent))
            {
                throw new InvalidOperationException(
                    "The DVH PNG output directory could not be resolved.");
            }

            Directory.CreateDirectory(parent);
            // The workspace uses native checkbox legends and deliberately hides
            // OxyPlot's legend. A standalone PNG must carry its own curve labels
            // and synthetic-use notice, without changing that GUI configuration.
            bool legendVisible = plotModel.IsLegendVisible;
            var legendPlacement = plotModel.LegendPlacement;
            var legendPosition = plotModel.LegendPosition;
            double legendFontSize = plotModel.LegendFontSize;
            double legendMaxWidth = plotModel.LegendMaxWidth;
            string title = plotModel.Title;
            double titleFontSize = plotModel.TitleFontSize;
            var legendEntries = plotModel.Series.ToDictionary(series => series, series => series.RenderInLegend);
            try
            {
                plotModel.IsLegendVisible = true;
                plotModel.LegendPlacement = LegendPlacement.Outside;
                plotModel.LegendPosition = LegendPosition.RightTop;
                plotModel.LegendFontSize = 14;
                plotModel.LegendMaxWidth = 280;
                plotModel.Title = "DVH - SYNTHETIC DATA - NOT FOR CLINICAL USE";
                plotModel.TitleFontSize = 20;
                foreach (var series in plotModel.Series) series.RenderInLegend = series.IsVisible;
                plotModel.InvalidatePlot(true);
                PngExporter.Export(plotModel, fullPath, ExportWidth, ExportHeight, OxyColors.White, 96.0);
            }
            finally
            {
                plotModel.IsLegendVisible = legendVisible;
                plotModel.LegendPlacement = legendPlacement;
                plotModel.LegendPosition = legendPosition;
                plotModel.LegendFontSize = legendFontSize;
                plotModel.LegendMaxWidth = legendMaxWidth;
                plotModel.Title = title;
                plotModel.TitleFontSize = titleFontSize;
                foreach (var entry in legendEntries) entry.Key.RenderInLegend = entry.Value;
                plotModel.InvalidatePlot(false);
            }
            return fullPath;
        }
    }
}
