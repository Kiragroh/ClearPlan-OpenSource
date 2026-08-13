using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;

namespace ClearPlan.Simulator
{
    public sealed class SimulatorScenarioRepository
    {
        private readonly string scenarioDirectory;

        public SimulatorScenarioRepository(string executableDirectory)
        {
            if (string.IsNullOrWhiteSpace(executableDirectory))
            {
                throw new ArgumentException(
                    "The executable directory is required.",
                    "executableDirectory");
            }

            scenarioDirectory = Path.Combine(
                Path.GetFullPath(executableDirectory),
                "Scenarios");
        }

        public IList<string> ScenarioIds
        {
            get
            {
                return new List<string>(
                    SyntheticScenarioFactory.ScenarioIds);
            }
        }

        public ReviewSnapshot Load(string scenarioId)
        {
            if (!SyntheticScenarioFactory.ScenarioIds.Contains(
                scenarioId,
                StringComparer.Ordinal))
            {
                throw new ArgumentException(
                    "Unknown checked-in simulator scenario: " + scenarioId,
                    "scenarioId");
            }

            string path = Path.Combine(
                scenarioDirectory,
                scenarioId + ".json");
            if (!File.Exists(path))
            {
                throw new FileNotFoundException(
                    "The checked-in simulator scenario was not copied.",
                    path);
            }

            ReviewSnapshot snapshot = ReviewSnapshotJson.DeserializeUtf8(
                File.ReadAllBytes(path));
            ReviewSnapshotValidationResult result =
                ReviewSnapshotValidator.ValidateForSimulator(snapshot);
            if (!result.IsValid)
            {
                string messages = string.Join(
                    Environment.NewLine,
                    result.Issues.Select(
                        issue =>
                            issue.Code + " " + issue.Path + ": " +
                            issue.Message));
                throw new InvalidDataException(
                    "Simulator scenario validation failed." +
                    Environment.NewLine + messages);
            }
            if (!string.Equals(
                snapshot.ScenarioId,
                scenarioId,
                StringComparison.Ordinal))
            {
                throw new InvalidDataException(
                    "Scenario filename and scenarioId do not match.");
            }

            return snapshot;
        }
    }
}
