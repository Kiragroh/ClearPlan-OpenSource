using System;
using System.IO;
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
            plotModel.InvalidatePlot(true);
            PngExporter.Export(
                plotModel,
                fullPath,
                ExportWidth,
                ExportHeight,
                OxyColors.White,
                96.0);
            return fullPath;
        }
    }
}
