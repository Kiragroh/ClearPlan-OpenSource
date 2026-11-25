using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClearPlan.Core.Collision;
using ClearPlan.Core.Review;
using Newtonsoft.Json;

namespace ClearPlan.Core.Simulation
{
    /// <summary>
    /// Sandbox from a real test plan (2026-11-20, Max): a read-only export of a Varian test patient plan
    /// (zzVarianCases) with placeholder identities and blurred MU, used instead of synthetic data so that
    /// zoom, BEV and collision look realistic. The file is marked synthetic: no ARIA upload, demo labels.
    /// </summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class SandboxSnapshotFile
    {
        public const string ScenarioIdValue = "sandbox-real-plan";
        public const string DefaultRelativePath = "Sandbox\\sandbox-snapshot.json";
        /// <summary>Only Varian test patients may be exported into a sandbox.</summary>
        public static readonly string[] AllowedPatientIds = { "zzVarianCases" };

        [JsonProperty] public int SchemaVersion { get; set; }
        [JsonProperty] public string CreatedUtc { get; set; }
        [JsonProperty] public string SourceDescription { get; set; }
        [JsonProperty] public double MuBlurMaximumFraction { get; set; }
        [JsonProperty] public string SnapshotJson { get; set; }
        [JsonProperty] public CollisionScene CollisionScene { get; set; }
        /// <summary>Control-point geometry that the review JSON deliberately omits (apertures, target
        /// projections, isocentres, field-start DRR pixels); restored on load.</summary>
        [JsonProperty] public List<SandboxControlPointGeometry> Geometry { get; set; }

        public static bool IsAllowedPatient(string patientId)
        {
            return AllowedPatientIds.Any(id => string.Equals(id, patientId, StringComparison.Ordinal));
        }

        /// <summary>Replaces identities, blurs every MU value of the plan by one deterministic factor per
        /// beam (within ±maximumFraction) and marks the snapshot as the sandbox scenario.</summary>
        public static void Sanitize(ReviewSnapshot snapshot, CollisionScene scene, string sourceDescription, double maximumFraction, int seed)
        {
            if (snapshot == null) throw new ArgumentNullException("snapshot");
            if (maximumFraction < 0 || maximumFraction > 0.5) throw new ArgumentOutOfRangeException("maximumFraction");
            snapshot.Synthetic = true;
            snapshot.ScenarioId = ScenarioIdValue;
            snapshot.ScenarioTitle = "Sandbox · Varian-Testplan";
            snapshot.ScenarioDescription = sourceDescription + " · MU unscharf (±" + Math.Round(maximumFraction * 100) + " %) · kein Patient, nicht klinisch.";
            snapshot.PatientDisplayLabel = "Sandbox · Varian-Testpatient";
            if (snapshot.Report != null)
            {
                snapshot.Report.ModeLabel = "SANDBOX · NICHT KLINISCH";
                snapshot.Report.Watermark = "SANDBOX";
                snapshot.Report.Subtitle = "Varian-Testplan · MU unscharf";
                snapshot.Report.OutputFileLabel = "clearplan-sandbox.pdf";
            }
            var random = new Random(seed);
            var analysis = snapshot.PlanAnalysis;
            if (analysis != null)
            {
                double totalBefore = analysis.TotalMetersetMu ?? 0, totalAfter = 0;
                foreach (var beam in analysis.Beams ?? new List<PlanAnalysis.ReviewBeamAnalysis>())
                {
                    double factor = 1 + (random.NextDouble() * 2 - 1) * maximumFraction;
                    if (beam.MetersetMu.HasValue) { beam.MetersetMu = Math.Round(beam.MetersetMu.Value * factor, 1); totalAfter += beam.MetersetMu.Value; }
                    foreach (var cp in beam.ControlPoints ?? new List<PlanAnalysis.ReviewControlPointSample>())
                    {
                        if (cp.IncrementalMetersetMu.HasValue) cp.IncrementalMetersetMu = cp.IncrementalMetersetMu.Value * factor;
                        if (cp.MetricMetersetWeightMu.HasValue) cp.MetricMetersetWeightMu = cp.MetricMetersetWeightMu.Value * factor;
                        if (cp.EstimatedSegmentDurationSeconds.HasValue) cp.EstimatedSegmentDurationSeconds = cp.EstimatedSegmentDurationSeconds.Value * factor;
                    }
                    if (beam.EstimatedBeamDurationSeconds.HasValue) beam.EstimatedBeamDurationSeconds = beam.EstimatedBeamDurationSeconds.Value * factor;
                }
                if (analysis.TotalMetersetMu.HasValue)
                {
                    double ratio = totalBefore > 0 && totalAfter > 0 ? totalAfter / totalBefore : 1 + (random.NextDouble() * 2 - 1) * maximumFraction;
                    analysis.TotalMetersetMu = Math.Round(totalBefore * ratio, 1);
                    if (analysis.MuPerGy.HasValue) analysis.MuPerGy = Math.Round(analysis.MuPerGy.Value * ratio, 1);
                }
            }
            // The validator keeps clinical and synthetic pixels apart: the sandbox carries one mode.
            foreach (var beam in (analysis == null ? null : analysis.Beams) ?? new List<PlanAnalysis.ReviewBeamAnalysis>())
                foreach (var cp in beam.ControlPoints ?? new List<PlanAnalysis.ReviewControlPointSample>())
                    if (cp.BevImage != null) cp.BevImage.Synthetic = true;
            foreach (var image in snapshot.PlanImages ?? new List<ReviewPlanImage>())
                if (image != null) image.Synthetic = true;
            if (scene != null) scene.Synthetic = true;
        }

        public static void Save(string path, ReviewSnapshot snapshot, CollisionScene scene, string sourceDescription, double maximumFraction)
        {
            var file = new SandboxSnapshotFile
            {
                SchemaVersion = 1,
                CreatedUtc = DateTime.UtcNow.ToString("O"),
                SourceDescription = sourceDescription,
                MuBlurMaximumFraction = maximumFraction,
                SnapshotJson = ReviewSnapshotJson.Serialize(snapshot),
                CollisionScene = scene,
                Geometry = CaptureGeometry(snapshot)
            };
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            File.WriteAllText(temp, JsonConvert.SerializeObject(file), new UTF8Encoding(false));
            if (File.Exists(path)) File.Replace(temp, path, path + ".previous");
            else File.Move(temp, path);
        }

        private static List<SandboxControlPointGeometry> CaptureGeometry(ReviewSnapshot snapshot)
        {
            var rows = new List<SandboxControlPointGeometry>();
            var beams = snapshot.PlanAnalysis == null ? null : snapshot.PlanAnalysis.Beams;
            if (beams == null) return rows;
            for (int b = 0; b < beams.Count; b++)
                for (int c = 0; c < (beams[b].ControlPoints ?? new List<PlanAnalysis.ReviewControlPointSample>()).Count; c++)
                {
                    var cp = beams[b].ControlPoints[c];
                    rows.Add(new SandboxControlPointGeometry
                    {
                        BeamIndex = b, ControlPointPosition = c, Aperture = cp.Aperture, TargetOutlines = cp.TargetOutlines,
                        TargetProjectionStrips = cp.TargetProjectionStrips, IsocenterMm = cp.IsocenterMm,
                        BevImage = cp.BevImage, BevPixels = cp.BevImage == null ? null : cp.BevImage.GrayscalePixels
                    });
                }
            return rows;
        }

        private static void RestoreGeometry(ReviewSnapshot snapshot, List<SandboxControlPointGeometry> rows)
        {
            var beams = snapshot.PlanAnalysis == null ? null : snapshot.PlanAnalysis.Beams;
            if (beams == null || rows == null) return;
            foreach (var row in rows)
            {
                if (row == null || row.BeamIndex < 0 || row.BeamIndex >= beams.Count) continue;
                var points = beams[row.BeamIndex].ControlPoints;
                if (points == null || row.ControlPointPosition < 0 || row.ControlPointPosition >= points.Count) continue;
                var cp = points[row.ControlPointPosition];
                cp.Aperture = row.Aperture;
                if (row.TargetOutlines != null) cp.TargetOutlines = row.TargetOutlines;
                if (row.TargetProjectionStrips != null) cp.TargetProjectionStrips = row.TargetProjectionStrips;
                cp.IsocenterMm = row.IsocenterMm;
                if (row.BevImage != null)
                {
                    row.BevImage.GrayscalePixels = row.BevPixels;
                    row.BevImage.Synthetic = true;
                    cp.BevImage = row.BevImage;
                }
            }
        }

        /// <summary>Loads a sandbox file; returns null when the file is missing, invalid or not a sandbox.</summary>
        public static ReviewSnapshot TryLoad(string path, out CollisionScene scene, out string description)
        {
            scene = null; description = null;
            try
            {
                if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return null;
                var file = JsonConvert.DeserializeObject<SandboxSnapshotFile>(File.ReadAllText(path, Encoding.UTF8));
                if (file == null || file.SchemaVersion != 1 || string.IsNullOrWhiteSpace(file.SnapshotJson)) return null;
                var snapshot = ReviewSnapshotJson.Deserialize(file.SnapshotJson);
                if (snapshot == null || !snapshot.Synthetic || snapshot.ScenarioId != ScenarioIdValue) return null;
                RestoreGeometry(snapshot, file.Geometry);
                scene = file.CollisionScene;
                if (scene != null) scene.Synthetic = true;
                snapshot.SandboxCollisionScene = scene;
                description = file.SourceDescription;
                return snapshot;
            }
            catch (Exception) { scene = null; description = null; return null; }
        }
    }

    /// <summary>Per control point geometry stored beside the review JSON in a sandbox file.</summary>
    public sealed class SandboxControlPointGeometry
    {
        public int BeamIndex { get; set; }
        public int ControlPointPosition { get; set; }
        public PlanAnalysis.ApertureGeometry Aperture { get; set; }
        public List<List<PlanAnalysis.BeamPoint>> TargetOutlines { get; set; }
        public List<PlanAnalysis.ApertureRectangle> TargetProjectionStrips { get; set; }
        public double[] IsocenterMm { get; set; }
        public PlanAnalysis.BeamEyeViewImage BevImage { get; set; }
        public byte[] BevPixels { get; set; }
    }
}
