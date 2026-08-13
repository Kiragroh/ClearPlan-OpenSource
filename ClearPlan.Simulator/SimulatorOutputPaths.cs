using System;
using System.IO;

namespace ClearPlan.Simulator
{
    public static class SimulatorOutputPaths
    {
        public static string GetDefaultReportDirectory(
            string executableDirectory)
        {
            if (string.IsNullOrWhiteSpace(executableDirectory))
            {
                throw new ArgumentException(
                    "The executable directory is required.",
                    "executableDirectory");
            }

            SimulatorArguments.ValidateLocalOutputDirectory(
                executableDirectory,
                "Simulator executable directory");
            string fullDirectory = Path.GetFullPath(executableDirectory);
            var directory = new DirectoryInfo(fullDirectory);
            DirectoryInfo simulatorDirectory = directory.Parent;
            DirectoryInfo artifactsDirectory =
                simulatorDirectory == null
                    ? null
                    : simulatorDirectory.Parent;

            if (simulatorDirectory != null &&
                artifactsDirectory != null &&
                simulatorDirectory.Name.Equals(
                    "simulator",
                    StringComparison.OrdinalIgnoreCase) &&
                artifactsDirectory.Name.Equals(
                    "artifacts",
                    StringComparison.OrdinalIgnoreCase))
            {
                return Path.Combine(
                    artifactsDirectory.FullName,
                    "simulator-reports");
            }

            return Path.Combine(
                fullDirectory,
                "artifacts",
                "simulator-reports");
        }

        public static string GetDefaultDvhExportPath(
            string reportDirectory,
            string scenarioId)
        {
            if (string.IsNullOrWhiteSpace(reportDirectory))
            {
                throw new ArgumentException(
                    "The report directory is required.",
                    "reportDirectory");
            }
            if (string.IsNullOrWhiteSpace(scenarioId) ||
                scenarioId.IndexOfAny(
                    Path.GetInvalidFileNameChars()) >= 0)
            {
                throw new ArgumentException(
                    "A file-safe scenario ID is required.",
                    "scenarioId");
            }

            SimulatorArguments.ValidateLocalOutputDirectory(
                reportDirectory,
                "DVH export directory");
            return Path.Combine(
                Path.GetFullPath(reportDirectory),
                scenarioId + "-dvh.png");
        }
    }
}
