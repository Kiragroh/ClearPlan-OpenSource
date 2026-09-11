using System;
using ClearPlan.Simulator;

namespace ClearPlan.Core.Tests
{
    internal static class SimulatorArgumentsTests
    {
        public static void UsesDocumentedDefaults()
        {
            SimulatorArguments arguments =
                SimulatorArguments.Parse(new string[0]);

            TestAssert.Equal("baseline-pass", arguments.ScenarioId);
            TestAssert.Equal("overview", arguments.TabId);
            TestAssert.Equal(null, arguments.CapturePath);
            TestAssert.Equal(null, arguments.CaptureAllDirectory);
            TestAssert.Equal(null, arguments.ReportPath);
            TestAssert.False(arguments.IsCaptureMode);
        }

        public static void ParsesScenarioAndTab()
        {
            SimulatorArguments arguments = SimulatorArguments.Parse(
                new[]
                {
                    "--scenario",
                    "field-and-mapping",
                    "--tab",
                    "fields"
                });

            TestAssert.Equal(
                "field-and-mapping",
                arguments.ScenarioId);
            TestAssert.Equal("fields", arguments.TabId);
            TestAssert.Equal("warnings", SimulatorArguments.Parse(new[] { "--tab", "warnings" }).TabId);
        }

        public static void ParsesSingleCapture()
        {
            SimulatorArguments arguments = SimulatorArguments.Parse(
                new[]
                {
                    "--tab",
                    "dvh",
                    "--capture",
                    @".\output\dvh.png",
                    "--report",
                    @".\output\report.pdf"
                });

            TestAssert.Equal("dvh", arguments.TabId);
            TestAssert.Equal(
                @".\output\dvh.png",
                arguments.CapturePath);
            TestAssert.Equal(
                @".\output\report.pdf",
                arguments.ReportPath);
            TestAssert.True(arguments.IsCaptureMode);
        }

        public static void ParsesCaptureAll()
        {
            SimulatorArguments arguments = SimulatorArguments.Parse(
                new[]
                {
                    "--scenario",
                    "metadata-plancheck",
                    "--capture-all",
                    @".\captures"
                });

            TestAssert.Equal(
                "metadata-plancheck",
                arguments.ScenarioId);
            TestAssert.Equal(@".\captures", arguments.CaptureAllDirectory);
            TestAssert.True(arguments.IsCaptureMode);
        }

        public static void ParsesLayoutProbe()
        {
            SimulatorArguments arguments = SimulatorArguments.Parse(
                new[]
                {
                    "--scenario",
                    "field-and-mapping",
                    "--layout-probe",
                    "1256.72x720",
                    "--output",
                    @".\layout\probe.json"
                });

            TestAssert.Equal(
                1256.72,
                arguments.LayoutProbeWidth);
            TestAssert.Equal(
                720.0,
                arguments.LayoutProbeHeight);
            TestAssert.Equal(
                @".\layout\probe.json",
                arguments.LayoutProbeOutputPath);
            TestAssert.True(arguments.IsLayoutProbeMode);
            TestAssert.False(arguments.IsCaptureMode);
        }

        public static void RejectsInvalidLayoutProbeRequests()
        {
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "1180x720"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--output",
                        "probe.json"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "1179x720",
                        "--output",
                        "probe.json"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "1180x719",
                        "--output",
                        "probe.json"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "not-a-size",
                        "--output",
                        "probe.json"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "1180x720",
                        "--output",
                        "probe.txt"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "1180x720",
                        "--output",
                        @"\\server\share\probe.json"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--layout-probe",
                        "1180x720",
                        "--output",
                        "probe.json",
                        "--capture",
                        "capture.png"
                    }));
        }

        public static void RejectsUnknownOrIncompleteValues()
        {
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--scenario", "not-a-scenario" }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--tab", "dose" }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--scenario" }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--unknown", "value" }));
        }

        public static void RejectsConflictingOrDuplicateValues()
        {
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--capture",
                        "one.png",
                        "--capture-all",
                        "all"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--tab",
                        "pqm",
                        "--capture-all",
                        "all"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--scenario",
                        "baseline-pass",
                        "--scenario",
                        "oar-overdose"
                    }));
        }

        public static void RejectsUnsafeOrWrongOutputPaths()
        {
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--capture", @"\\server\share\shot.png" }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[]
                    {
                        "--capture",
                        "file://server/share/shot.png"
                    }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--capture", "shot.jpg" }));
            TestAssert.Throws<ArgumentException>(
                () => SimulatorArguments.Parse(
                    new[] { "--report", "report.txt" }));
        }
    }
}
