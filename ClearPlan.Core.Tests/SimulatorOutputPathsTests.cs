using System;
using System.IO;
using ClearPlan.Simulator;

namespace ClearPlan.Core.Tests
{
    internal static class SimulatorOutputPathsTests
    {
        public static void UsesRepositoryArtifactsReportDirectory()
        {
            string repositoryRoot = Path.Combine(
                Path.GetPathRoot(Environment.CurrentDirectory),
                "clearplan-repository");
            string executableDirectory = Path.Combine(
                repositoryRoot,
                "artifacts",
                "simulator",
                "Release");

            string actual =
                SimulatorOutputPaths.GetDefaultReportDirectory(
                    executableDirectory);

            TestAssert.Equal(
                Path.Combine(
                    repositoryRoot,
                    "artifacts",
                    "simulator-reports"),
                actual);
        }

        public static void UsesPortableArtifactsReportDirectory()
        {
            string portableRoot = Path.Combine(
                Path.GetPathRoot(Environment.CurrentDirectory),
                "clearplan-portable");

            string actual =
                SimulatorOutputPaths.GetDefaultReportDirectory(
                    portableRoot);

            TestAssert.Equal(
                Path.Combine(
                    portableRoot,
                    "artifacts",
                    "simulator-reports"),
                actual);
        }

        public static void RejectsNetworkExecutableDirectory()
        {
            TestAssert.Throws<ArgumentException>(
                () => SimulatorOutputPaths.GetDefaultReportDirectory(
                    @"\\server\share\simulator"));
        }

        public static void UsesDeterministicLocalDvhExportPath()
        {
            string reportDirectory = Path.Combine(
                Path.GetPathRoot(Environment.CurrentDirectory),
                "clearplan-repository",
                "artifacts",
                "simulator-reports");

            string actual =
                SimulatorOutputPaths.GetDefaultDvhExportPath(
                    reportDirectory,
                    "baseline-pass");

            TestAssert.Equal(
                Path.Combine(
                    reportDirectory,
                    "baseline-pass-dvh.png"),
                actual);
        }
    }
}
