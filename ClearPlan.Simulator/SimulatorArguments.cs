using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;

namespace ClearPlan.Simulator
{
    public sealed class SimulatorArguments
    {
        public const string SyntheticMarker =
            "SYNTHETIC DEMONSTRATION";

        public const string ClinicalUseMarker =
            "NOT FOR CLINICAL USE";

        public const string SafetyBanner =
            SyntheticMarker + " — " + ClinicalUseMarker;

        private static readonly string[] AllowedTabs =
        {
            "overview",
            "pqm",
            "plancheck",
            "fields",
            "dvh",
            "images",
            "collision",
            "parameters",
            "comparison",
            "bev",
            "warnings"
        };

        private SimulatorArguments()
        {
            ScenarioId = "baseline-pass";
            TabId = "overview";
        }

        public string ScenarioId { get; private set; }

        public string TabId { get; private set; }

        public string CapturePath { get; private set; }

        public string CaptureAllDirectory { get; private set; }

        public string ReportPath { get; private set; }

        public double LayoutProbeWidth { get; private set; }

        public double LayoutProbeHeight { get; private set; }

        public string LayoutProbeOutputPath { get; private set; }

        public bool IsCaptureMode
        {
            get
            {
                return !string.IsNullOrWhiteSpace(CapturePath) ||
                       !string.IsNullOrWhiteSpace(CaptureAllDirectory);
            }
        }

        public bool IsLayoutProbeMode
        {
            get
            {
                return LayoutProbeWidth > 0.0 &&
                       LayoutProbeHeight > 0.0 &&
                       !string.IsNullOrWhiteSpace(
                           LayoutProbeOutputPath);
            }
        }

        public bool IsAutomatedMode
        {
            get
            {
                return IsCaptureMode || IsLayoutProbeMode;
            }
        }

        public static SimulatorArguments Parse(string[] commandLine)
        {
            string[] values = commandLine ?? new string[0];
            var parsed = new SimulatorArguments();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool tabSpecified = false;

            for (int index = 0; index < values.Length; index++)
            {
                string option = values[index] ?? string.Empty;
                if (!IsKnownOption(option))
                {
                    throw new ArgumentException(
                        "Unknown simulator option: " + option,
                        "commandLine");
                }
                if (!seen.Add(option))
                {
                    throw new ArgumentException(
                        "Simulator option was provided more than once: " +
                        option,
                        "commandLine");
                }
                if (index + 1 >= values.Length ||
                    string.IsNullOrWhiteSpace(values[index + 1]) ||
                    values[index + 1].StartsWith(
                        "--",
                        StringComparison.Ordinal))
                {
                    throw new ArgumentException(
                        "Simulator option requires a value: " + option,
                        "commandLine");
                }

                string optionValue = values[++index];
                switch (option.ToLowerInvariant())
                {
                    case "--scenario":
                        parsed.ScenarioId = optionValue;
                        break;
                    case "--tab":
                        parsed.TabId = optionValue;
                        tabSpecified = true;
                        break;
                    case "--capture":
                        parsed.CapturePath = optionValue;
                        break;
                    case "--capture-all":
                        parsed.CaptureAllDirectory = optionValue;
                        break;
                    case "--report":
                        parsed.ReportPath = optionValue;
                        break;
                    case "--layout-probe":
                        double layoutProbeWidth;
                        double layoutProbeHeight;
                        ParseLayoutProbeSize(
                            optionValue,
                            out layoutProbeWidth,
                            out layoutProbeHeight);
                        parsed.LayoutProbeWidth = layoutProbeWidth;
                        parsed.LayoutProbeHeight = layoutProbeHeight;
                        break;
                    case "--output":
                        parsed.LayoutProbeOutputPath = optionValue;
                        break;
                }
            }

            if (!SyntheticScenarioFactory.ScenarioIds.Contains(
                parsed.ScenarioId,
                StringComparer.Ordinal) && !SyntheticPublicationScenarioFactory.ScenarioIds.Contains(parsed.ScenarioId, StringComparer.Ordinal))
            {
                throw new ArgumentException(
                    "Unknown synthetic scenario: " + parsed.ScenarioId,
                    "commandLine");
            }
            if (!AllowedTabs.Contains(
                parsed.TabId,
                StringComparer.Ordinal))
            {
                throw new ArgumentException(
                    "Unknown simulator tab: " + parsed.TabId,
                    "commandLine");
            }
            if (!string.IsNullOrWhiteSpace(parsed.CapturePath) &&
                !string.IsNullOrWhiteSpace(parsed.CaptureAllDirectory))
            {
                throw new ArgumentException(
                    "--capture and --capture-all cannot be combined.",
                    "commandLine");
            }
            if (tabSpecified &&
                !string.IsNullOrWhiteSpace(parsed.CaptureAllDirectory))
            {
                throw new ArgumentException(
                    "--tab cannot be combined with --capture-all.",
                    "commandLine");
            }

            bool layoutSizeSpecified =
                parsed.LayoutProbeWidth > 0.0 &&
                parsed.LayoutProbeHeight > 0.0;
            bool layoutOutputSpecified =
                !string.IsNullOrWhiteSpace(
                    parsed.LayoutProbeOutputPath);
            if (layoutSizeSpecified != layoutOutputSpecified)
            {
                throw new ArgumentException(
                    "--layout-probe and --output must be provided together.",
                    "commandLine");
            }
            if (layoutSizeSpecified &&
                (parsed.IsCaptureMode ||
                 !string.IsNullOrWhiteSpace(parsed.ReportPath)))
            {
                throw new ArgumentException(
                    "Layout probing cannot be combined with capture or report output.",
                    "commandLine");
            }

            ValidateLocalOutputFile(
                parsed.CapturePath,
                ".png",
                "--capture");
            ValidateLocalOutputDirectory(
                parsed.CaptureAllDirectory,
                "--capture-all");
            ValidateLocalOutputFile(
                parsed.ReportPath,
                ".pdf",
                "--report");
            ValidateLocalOutputFile(
                parsed.LayoutProbeOutputPath,
                ".json",
                "--output");
            return parsed;
        }

        private static void ParseLayoutProbeSize(
            string value,
            out double width,
            out double height)
        {
            width = 0.0;
            height = 0.0;
            string size = value ?? string.Empty;
            int separatorIndex = size.IndexOf(
                'x');
            if (separatorIndex < 0)
            {
                separatorIndex = size.IndexOf('X');
            }
            if (separatorIndex <= 0 ||
                separatorIndex >= size.Length - 1 ||
                size.IndexOf(
                    'x',
                    separatorIndex + 1) >= 0 ||
                size.IndexOf(
                    'X',
                    separatorIndex + 1) >= 0 ||
                !double.TryParse(
                    size.Substring(0, separatorIndex),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out width) ||
                !double.TryParse(
                    size.Substring(separatorIndex + 1),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out height) ||
                double.IsNaN(width) ||
                double.IsInfinity(width) ||
                double.IsNaN(height) ||
                double.IsInfinity(height) ||
                width < ReviewWindowSizePolicy.MinimumWidth ||
                height < ReviewWindowSizePolicy.MinimumHeight)
            {
                throw new ArgumentException(
                    string.Format(
                        CultureInfo.InvariantCulture,
                        "--layout-probe requires WIDTHxHEIGHT with at least {0}x{1}.",
                        ReviewWindowSizePolicy.MinimumWidth,
                        ReviewWindowSizePolicy.MinimumHeight),
                    "value");
            }
        }

        internal static void ValidateLocalOutputFile(
            string path,
            string requiredExtension,
            string optionName)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }

            ValidateLocalOutputDirectory(path, optionName);
            if (!string.Equals(
                Path.GetExtension(path),
                requiredExtension,
                StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    optionName + " requires a " +
                    requiredExtension + " output path.",
                    "path");
            }
        }

        internal static void ValidateLocalOutputDirectory(
            string path,
            string optionName)
        {
            if (string.IsNullOrWhiteSpace(path))
            {
                return;
            }
            if (path.StartsWith(@"\\", StringComparison.Ordinal) ||
                path.StartsWith("//", StringComparison.Ordinal))
            {
                throw new ArgumentException(
                    optionName + " accepts local output paths only.",
                    "path");
            }

            Uri uri;
            if (Uri.TryCreate(path, UriKind.Absolute, out uri) &&
                (!uri.IsFile || uri.IsUnc ||
                 !string.IsNullOrWhiteSpace(uri.Host)))
            {
                throw new ArgumentException(
                    optionName + " accepts local file paths only.",
                    "path");
            }
        }

        private static bool IsKnownOption(string option)
        {
            return option.Equals(
                       "--scenario",
                       StringComparison.OrdinalIgnoreCase) ||
                   option.Equals(
                       "--tab",
                       StringComparison.OrdinalIgnoreCase) ||
                   option.Equals(
                       "--capture",
                       StringComparison.OrdinalIgnoreCase) ||
                   option.Equals(
                       "--capture-all",
                       StringComparison.OrdinalIgnoreCase) ||
                   option.Equals(
                       "--report",
                       StringComparison.OrdinalIgnoreCase) ||
                   option.Equals(
                       "--layout-probe",
                       StringComparison.OrdinalIgnoreCase) ||
                   option.Equals(
                       "--output",
                       StringComparison.OrdinalIgnoreCase);
        }
    }
}
