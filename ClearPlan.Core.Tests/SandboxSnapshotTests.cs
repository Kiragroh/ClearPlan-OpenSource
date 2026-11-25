using System;
using System.IO;
using System.Linq;
using ClearPlan.Core.Collision;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;

namespace ClearPlan.Core.Tests
{
    /// <summary>Sandbox from a real test plan (2026-11-20): allowlist, MU blur, identities, file round trip.</summary>
    internal static class SandboxSnapshotTests
    {
        public static void SanitizeAndRoundTrip()
        {
            TestAssert.True(SandboxSnapshotFile.IsAllowedPatient("zzVarianCases"));
            TestAssert.False(SandboxSnapshotFile.IsAllowedPatient("12345678"), "Only the Varian test patient may be exported.");
            TestAssert.False(SandboxSnapshotFile.IsAllowedPatient("zzvariancases"), "The allowlist is exact.");

            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            var scene = SyntheticCollisionFactory.Create(snapshot.ActivePlanKey);
            var beamsBefore = snapshot.PlanAnalysis.Beams.Select(b => b.MetersetMu).ToList();
            SandboxSnapshotFile.Sanitize(snapshot, scene, "Varian-Testpatient · Kurs X · Plan Y", 0.08, 7);
            TestAssert.True(snapshot.Synthetic && snapshot.ScenarioId == SandboxSnapshotFile.ScenarioIdValue);
            TestAssert.Equal("Sandbox · Varian-Testpatient", snapshot.PatientDisplayLabel);
            for (int i = 0; i < beamsBefore.Count; i++)
            {
                if (!beamsBefore[i].HasValue) continue;
                double ratio = snapshot.PlanAnalysis.Beams[i].MetersetMu.Value / beamsBefore[i].Value;
                TestAssert.True(ratio >= 0.919 && ratio <= 1.081, "MU blur must stay within ±8 %, ratio " + ratio);
            }
            TestAssert.True(ReviewSnapshotValidator.Validate(snapshot).IsValid, "The sanitized sandbox must pass snapshot validation.");

            string dir = Path.Combine(Path.GetTempPath(), "clearplan-sandbox-" + Guid.NewGuid().ToString("N"));
            try
            {
                string path = Path.Combine(dir, SandboxSnapshotFile.DefaultRelativePath);
                SandboxSnapshotFile.Save(path, snapshot, scene, "Varian-Testpatient · Kurs X · Plan Y", 0.08);
                CollisionScene loadedScene; string description;
                var loaded = SandboxSnapshotFile.TryLoad(path, out loadedScene, out description);
                TestAssert.NotNull(loaded, "Saved sandbox must load.");
                TestAssert.True(loadedScene != null && loadedScene.Surfaces.Count == scene.Surfaces.Count && loadedScene.Poses.Count == scene.Poses.Count);
                TestAssert.True(ReferenceEquals(loaded.SandboxCollisionScene, loadedScene));
                TestAssert.Equal("Varian-Testpatient · Kurs X · Plan Y", description);
                var cpBefore = snapshot.PlanAnalysis.Beams.SelectMany(b => b.ControlPoints).FirstOrDefault(c => c.Aperture != null);
                if (cpBefore != null)
                {
                    var cpAfter = loaded.PlanAnalysis.Beams.SelectMany(b => b.ControlPoints).First(c => c.Index == cpBefore.Index);
                    TestAssert.NotNull(cpAfter.Aperture, "MLC geometry must survive the sandbox file.");
                    TestAssert.Equal(cpBefore.Aperture.EffectiveOpenings.Count, cpAfter.Aperture.EffectiveOpenings.Count);
                }
                TestAssert.Equal("SANDBOX · NICHT KLINISCH", loaded.Report.ModeLabel);
                File.WriteAllText(path, "{ \"SchemaVersion\": 1, \"SnapshotJson\": \"{}\" }");
                TestAssert.True(SandboxSnapshotFile.TryLoad(path, out loadedScene, out description) == null, "Invalid files fall back to the synthetic sandbox.");
            }
            finally { try { Directory.Delete(dir, true); } catch (Exception) { } }
        }
    }
}
