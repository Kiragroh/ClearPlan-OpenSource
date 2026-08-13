using System;
using System.IO;
using ClearPlan.Core.Review;
using ClearPlan.Reporting;
using ClearPlan.Reporting.MigraDoc;

namespace ClearPlan.Simulator
{
    public sealed class SyntheticReportService
    {
        public string Export(ReviewSnapshot snapshot, string outputPath)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException("snapshot");
            }
            ReviewSnapshotValidationResult validation =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);
            if (!validation.IsValid)
            {
                throw new InvalidDataException(
                    "A valid synthetic snapshot is required for reporting.");
            }

            SimulatorArguments.ValidateLocalOutputFile(
                outputPath,
                ".pdf",
                "PDF report");
            string fullPath = Path.GetFullPath(outputPath);
            string parent = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(parent))
            {
                throw new InvalidOperationException(
                    "The PDF output directory could not be resolved.");
            }

            Directory.CreateDirectory(parent);
            ReviewReportDocument document =
                new ReviewSnapshotReportMapper().Map(snapshot);
            document.ModeLabel = SimulatorArguments.SafetyBanner;
            document.Watermark = SimulatorArguments.SafetyBanner;
            new ReportPdf().Export(fullPath, document);
            return fullPath;
        }
    }
}
