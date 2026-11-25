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

        private string SandboxPath { get { return Path.Combine(Path.GetDirectoryName(scenarioDirectory), SandboxSnapshotFile.DefaultRelativePath); } }

        public IList<string> ScenarioIds
        {
            get
            {
                var ids = new List<string>();
                if (File.Exists(SandboxPath)) ids.Add(SandboxSnapshotFile.ScenarioIdValue);
                ids.AddRange(SyntheticScenarioFactory.ScenarioIds.Concat(SyntheticPublicationScenarioFactory.ScenarioIds));
                return ids;
            }
        }

        public ReviewSnapshot Load(string scenarioId)
        {
            // 2026-11-20: exported Varian test plan (MU blurred) next to the simulator, if present.
            if (string.Equals(scenarioId, SandboxSnapshotFile.ScenarioIdValue, StringComparison.Ordinal))
            {
                Core.Collision.CollisionScene scene; string description;
                var sandbox = SandboxSnapshotFile.TryLoad(SandboxPath, out scene, out description);
                if (sandbox == null) throw new FileNotFoundException("The sandbox snapshot is missing or invalid.", SandboxPath);
                return sandbox;
            }
            // Publication geometry is generated in memory; JSON intentionally omits MLC/DRR pixels.
            if (SyntheticPublicationScenarioFactory.ScenarioIds.Contains(scenarioId, StringComparer.Ordinal))
                return SyntheticPublicationScenarioFactory.Create(scenarioId);
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

            // Rehydrate deterministic in-memory MLC geometry, intentionally omitted from portable JSON.
            SyntheticScenarioFactory.AddPlanAnalysis(snapshot);
            return snapshot;
        }
    }
}
