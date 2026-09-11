using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class DoseRateEstimatorTests
    {
        public static void ContractExists()
        {
            var assembly = typeof(ReviewBeamAnalysis).Assembly;
            TestAssert.NotNull(assembly.GetType("ClearPlan.Core.PlanAnalysis.DoseRateEstimator"),
                "Missing detached PlanCheck-style dose-rate estimator.");
            TestAssert.NotNull(assembly.GetType("ClearPlan.Core.PlanAnalysis.DoseRateEstimationProfileCatalog"));
            foreach (string name in new[] { "MachineModelName", "MachineModel", "GantryDirection", "DoseRateEstimateStatus",
                "DoseRateEstimateReason", "DoseRateEstimateProfile", "EstimatedBeamDurationSeconds", "DoseRateEstimateMaxGantrySpeedDegreesPerSecond" })
                TestAssert.NotNull(typeof(ReviewBeamAnalysis).GetProperty(name), "Missing estimate metadata: " + name);
            foreach (string name in new[] { "EstimatedDoseRateMuPerMin", "EstimatedSegmentDurationSeconds" })
                TestAssert.Equal(typeof(double?), typeof(ReviewControlPointSample).GetProperty(name).PropertyType);
        }

        public static void GantryAndMuLimits()
        {
            foreach (double speed in new[] { 6d, 12d })
            {
                var beam = Beam(new[] { 0d, 12d, 18d }, new[] { 0d, 0.1d, 1d });
                DoseRateEstimator.Apply(beam, Profile(speed));
                TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
                Near(speed * 5, beam.ControlPoints[1].EstimatedDoseRateMuPerMin);
                Near(540, beam.ControlPoints[2].EstimatedDoseRateMuPerMin);
                Near(12 / speed, beam.ControlPoints[1].EstimatedSegmentDurationSeconds);
                Near(1, beam.ControlPoints[2].EstimatedSegmentDurationSeconds);
                Near(12 / speed + 1, beam.EstimatedBeamDurationSeconds);
                TestAssert.Equal("synthetic-profile", beam.DoseRateEstimateProfile);
                Near(speed, beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond);
                TestAssert.True(beam.DoseRateEstimateReason.Contains("ESTIMATE") && beam.DoseRateEstimateReason.Contains("PlanCheck"));
                foreach (string caveat in new[] { "acceleration", "MLC", "jaw", "ramping", "holds", "not time-resolved", "Synthetic test assumption" })
                    TestAssert.True(beam.DoseRateEstimateReason.Contains(caveat), "Missing estimate limitation: " + caveat);
            }
        }

        public static void ConstantMuLimitedTraceIsValid()
        {
            var beam = Beam(new[] { 0d, 1d, 2d }, new[] { 0d, 0.3d, 1d });
            beam.MetersetMu = 100;
            DoseRateEstimator.Apply(beam, Profile());
            TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
            Near(540, beam.ControlPoints[1].EstimatedDoseRateMuPerMin);
            Near(540, beam.ControlPoints[2].EstimatedDoseRateMuPerMin);
            Near(100d / 9d, beam.EstimatedBeamDurationSeconds);
        }

        public static void CmwNormalizationAndPlannedValuesStaySeparate()
        {
            var beam = Beam(new[] { 0d, 12d, 24d }, new[] { 0d, 1d, 10d });
            beam.ControlPoints[0].PlannedDoseRateMuPerMin = 111;
            beam.ControlPoints[1].PlannedDoseRateMuPerMin = null;
            beam.ControlPoints[2].PlannedDoseRateMuPerMin = 333;
            beam.ControlPoints[1].IncrementalMetersetMu = 999; // The estimator must derive its own normalized interval MU.
            DoseRateEstimator.Apply(beam, Profile());
            TestAssert.False(beam.ControlPoints[0].EstimatedDoseRateMuPerMin.HasValue);
            TestAssert.False(beam.ControlPoints[0].EstimatedSegmentDurationSeconds.HasValue);
            Near(30, beam.ControlPoints[1].EstimatedDoseRateMuPerMin);
            Near(270, beam.ControlPoints[2].EstimatedDoseRateMuPerMin);
            Near(111, beam.ControlPoints[0].PlannedDoseRateMuPerMin);
            TestAssert.False(beam.ControlPoints[1].PlannedDoseRateMuPerMin.HasValue);
            Near(333, beam.ControlPoints[2].PlannedDoseRateMuPerMin);
            Near(999, beam.ControlPoints[1].IncrementalMetersetMu);
            Near(540, beam.NominalDoseRateMuPerMin);
        }

        public static void WrapDirectionAndZeroIntervals()
        {
            foreach (string direction in new[] { "Clockwise", "CounterClockwise" })
            {
                var angles = direction == "Clockwise" ? new[] { 359d, 1d, 3d } : new[] { 1d, 359d, 357d };
                var beam = Beam(angles, new[] { 0d, 0.01d, 1d }); beam.GantryDirection = direction;
                DoseRateEstimator.Apply(beam, Profile());
                TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
                Near(18, beam.ControlPoints[1].EstimatedDoseRateMuPerMin);
                Near(1d / 3d, beam.ControlPoints[1].EstimatedSegmentDurationSeconds);
            }
            var zero = Beam(new[] { 0d, 12d, 12d, 24d }, new[] { 0d, 0d, 0d, 1d });
            DoseRateEstimator.Apply(zero, Profile());
            Near(0, zero.ControlPoints[1].EstimatedDoseRateMuPerMin);
            Near(2, zero.ControlPoints[1].EstimatedSegmentDurationSeconds);
            TestAssert.False(zero.ControlPoints[2].EstimatedDoseRateMuPerMin.HasValue);
            Near(0, zero.ControlPoints[2].EstimatedSegmentDurationSeconds);
            Near(300, zero.ControlPoints[3].EstimatedDoseRateMuPerMin);
            var endpoint = Beam(new[] { 359d, 360d, 0d, 1d }, new[] { 0d, 0d, 0d, 1d });
            DoseRateEstimator.Apply(endpoint, Profile());
            TestAssert.Equal("Estimated", endpoint.DoseRateEstimateStatus);
            Near(0, endpoint.ControlPoints[2].EstimatedSegmentDurationSeconds);
        }

        public static void InvalidInputsFailAtomically()
        {
            var mutations = new Action<ReviewBeamAnalysis>[]
            {
                b => b.MetersetMu = null, b => b.MetersetMu = 0, b => b.MetersetMu = -1,
                b => b.MetersetMu = double.NaN, b => b.MetersetMu = double.PositiveInfinity,
                b => b.NominalDoseRateMuPerMin = null, b => b.NominalDoseRateMuPerMin = 0,
                b => b.NominalDoseRateMuPerMin = -1, b => b.NominalDoseRateMuPerMin = double.NaN,
                b => b.NominalDoseRateMuPerMin = double.PositiveInfinity,
                b => b.Technique = null, b => b.Technique = "STATIC", b => b.Technique = "not VMAT",
                b => b.GantryDirection = null, b => b.GantryDirection = "None", b => b.GantryDirection = "CW",
                b => b.ControlPoints = null, b => b.ControlPoints.Clear(), b => b.ControlPoints.RemoveRange(1, 2),
                b => b.ControlPoints[0].CumulativeMetersetWeight = 0.01,
                b => b.ControlPoints[1].CumulativeMetersetWeight = -0.1,
                b => b.ControlPoints[1].CumulativeMetersetWeight = 2,
                b => b.ControlPoints[2].CumulativeMetersetWeight = double.NaN,
                b => b.ControlPoints[2].CumulativeMetersetWeight = double.PositiveInfinity,
                b => { foreach (var cp in b.ControlPoints) cp.CumulativeMetersetWeight = 0; },
                b => b.ControlPoints[2].GantryAngleDegrees = double.NaN,
                b => b.ControlPoints[2].GantryAngleDegrees = double.PositiveInfinity,
                b => b.ControlPoints[2].GantryAngleDegrees = -1,
                b => b.ControlPoints[2].GantryAngleDegrees = 361,
                b => b.ControlPoints[2].GantryAngleDegrees = 11, // Reversal, not a 359-degree interval.
                b => b.ControlPoints[2].GantryAngleDegrees = 192, // Ambiguous half-turn.
                b => b.ControlPoints[2].GantryAngleDegrees = 193,
                b => b.ControlPoints[0].Index = 1, b => b.ControlPoints[1].Index = -1,
                b => b.ControlPoints[1].Index = 2, b => b.ControlPoints[2].Index = 1,
                b => b.ControlPoints[2] = null
            };
            foreach (var mutate in mutations)
            {
                var beam = Beam(); SeedStale(beam); mutate(beam);
                DoseRateEstimator.Apply(beam, Profile());
                Unavailable(beam);
                if (beam.ControlPoints != null)
                    foreach (var cp in beam.ControlPoints.Where(p => p != null)) Near(123, cp.PlannedDoseRateMuPerMin);
            }
            foreach (string technique in new[] { "ARC", "VMAT", "RapidArc", "vmat" })
            {
                var beam = Beam(); beam.Technique = technique;
                DoseRateEstimator.Apply(beam, Profile());
                TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
            }
            TestAssert.Throws<ArgumentNullException>(() => DoseRateEstimator.Apply(null, Profile()));
        }

        public static void RepeatedApplyClearsStaleResults()
        {
            var beam = Beam(); DoseRateEstimator.Apply(beam, Profile());
            TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
            beam.ControlPoints[2].CumulativeMetersetWeight = -1;
            DoseRateEstimator.Apply(beam, Profile()); Unavailable(beam);
            beam.ControlPoints[2].CumulativeMetersetWeight = 1;
            DoseRateEstimator.Apply(beam, Profile());
            TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
            DoseRateEstimator.Apply(beam, (DoseRateEstimationProfile)null); Unavailable(beam);
            DoseRateEstimator.Apply(beam, Profile());
            DoseRateEstimator.Apply(beam, (DoseRateEstimationProfileCatalog)null); Unavailable(beam);
        }

        public static void FiniteOutputsAndOverflow()
        {
            var random = new Random(4183);
            foreach (double speed in new[] { 6d, 12d })
            for (int trial = 0; trial < 100; trial++)
            {
                double step = 0.01 + random.NextDouble() * 150;
                var beam = Beam(new[] { 0d, step, 2 * step }, new[] { 0d, random.NextDouble(), 1d });
                beam.MetersetMu = 1 + random.NextDouble() * 2000;
                beam.NominalDoseRateMuPerMin = 100 + random.NextDouble() * 2400;
                DoseRateEstimator.Apply(beam, Profile(speed));
                TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
                double reconstructedMu = 0;
                foreach (var cp in beam.ControlPoints.Skip(1))
                {
                    double rate = cp.EstimatedDoseRateMuPerMin.Value, time = cp.EstimatedSegmentDurationSeconds.Value;
                    TestAssert.True(!double.IsNaN(rate) && !double.IsInfinity(rate) && rate >= 0 && rate <= beam.NominalDoseRateMuPerMin.Value);
                    TestAssert.True(!double.IsNaN(time) && !double.IsInfinity(time) && time > 0);
                    reconstructedMu += rate * time / 60;
                }
                Near(beam.MetersetMu.Value, reconstructedMu);
            }
            var overflow = Beam(); overflow.MetersetMu = double.MaxValue; overflow.NominalDoseRateMuPerMin = 1;
            DoseRateEstimator.Apply(overflow, Profile()); Unavailable(overflow);
            var slow = Profile(double.Epsilon);
            var beamSlow = Beam(); DoseRateEstimator.Apply(beamSlow, slow); Unavailable(beamSlow);
        }

        public static void ProfileMatchingIsExactAndSpecific()
        {
            var shipped = DoseRateEstimationProfileCatalog.Parse(File.ReadAllText(Path.Combine("ClearPlan.Script", "Distribution", "MachineGeometry", "DoseRateProfiles.example.json")));
            string shippedReason;
            var nativeTds = Beam(); nativeTds.MachineModelName = "TDS"; nativeTds.MlcModel = "Varian High Definition 120";
            var nativeRds = Beam(); nativeRds.MachineModelName = "RDS"; nativeRds.MlcModel = "SX2";
            TestAssert.Equal("TrueBeam-HD120-PlanCheck-estimate", shipped.Resolve(nativeTds, out shippedReason).Id);
            TestAssert.Equal("Halcyon-SX2-RapidArc-estimate", shipped.Resolve(nativeRds, out shippedReason).Id);
            TestAssert.Equal("TDS", nativeTds.MachineModelName, "Matching must not rename the native machine.");
            nativeTds.MlcModel = "SX2"; nativeRds.MlcModel = "Varian High Definition 120";
            TestAssert.True(shipped.Resolve(nativeTds, out shippedReason) == null && shipped.Resolve(nativeRds, out shippedReason) == null,
                "Do not exchange distinct machine profiles based on a loose TDS/RDS or MLC alias.");
            var generic = Profile(); generic.MachineModelNames = new List<string> { "Synthetic model", "Second model" };
            var exact = Profile(12); exact.Id = "id-specific"; exact.MachineIds = new List<string> { "Synthetic-ID" };
            exact.MachineModels = new List<string> { "Synthetic enum" }; exact.MlcModels = new List<string> { "Synthetic MLC" };
            var catalog = Catalog(generic, exact); var beam = Beam(); string reason;
            TestAssert.Equal(exact, catalog.Resolve(beam, out reason));
            beam.MachineId = "synthetic-id"; beam.MlcModel = "synthetic mlc";
            TestAssert.Equal(exact, catalog.Resolve(beam, out reason));
            beam.MlcModel = "different";
            TestAssert.Equal(generic, catalog.Resolve(beam, out reason));
            beam.MachineModelName = "Second model";
            TestAssert.Equal(generic, catalog.Resolve(beam, out reason));
            foreach (string unknown in new[] { null, "", "Synthetic model extra", "prefix Synthetic model", "Synthetic model " })
            {
                beam.MachineModelName = unknown;
                TestAssert.True(catalog.Resolve(beam, out reason) == null);
                TestAssert.True(!string.IsNullOrWhiteSpace(reason));
            }
            var both = Profile(); both.MachineModels = new List<string> { "Synthetic enum" };
            beam = Beam(); beam.MachineModel = "different";
            TestAssert.True(Catalog(both).Resolve(beam, out reason) == null, "Model criteria must be AND-combined.");
            var idOnly = Profile(); idOnly.MachineModelNames = null; idOnly.MachineIds = new List<string> { "Synthetic-ID" };
            TestAssert.Equal(idOnly, Catalog(idOnly).Resolve(beam, out reason));
            var modelOnly = Profile(); modelOnly.MachineModelNames = null; modelOnly.MachineModels = new List<string> { "Synthetic enum" };
            TestAssert.Equal(modelOnly, Catalog(modelOnly).Resolve(Beam(), out reason));
        }

        public static void MissingAmbiguousInvalidProfilesAreUnavailable()
        {
            string reason; var beam = Beam();
            var first = Profile(); var second = Profile(12); second.Id = "second";
            var ambiguous = Catalog(first, second);
            TestAssert.True(ambiguous.Resolve(beam, out reason) == null && reason.Contains("ambiguous"));
            SeedStale(beam); DoseRateEstimator.Apply(beam, ambiguous); Unavailable(beam);
            first.MachineIds = new List<string> { beam.MachineId }; second.MachineIds = new List<string> { beam.MachineId };
            TestAssert.True(ambiguous.Resolve(beam, out reason) == null && reason.Contains("ambiguous"));
            var valid = Catalog(Profile()); valid.Profiles[0].MaxGantrySpeedDegreesPerSecond = double.NaN;
            TestAssert.True(valid.Resolve(beam, out reason) == null && reason.Contains("invalid"));
            DoseRateEstimator.Apply(beam, valid); Unavailable(beam);
            DoseRateEstimator.Apply(beam, Catalog()); Unavailable(beam);
            beam.MachineModelName = "Unknown machine";
            DoseRateEstimator.Apply(beam, Catalog(Profile())); Unavailable(beam);
            DoseRateEstimator.Apply(beam, Profile()); Unavailable(beam);
        }

        public static void ProfileJsonValidationIsStrictAndSanitized()
        {
            var catalog = Catalog(Profile());
            var parsed = DoseRateEstimationProfileCatalog.Parse(JsonConvert.SerializeObject(catalog));
            TestAssert.Equal(0, parsed.Validate().Count);
            TestAssert.Equal("synthetic-profile", parsed.Profiles[0].Id);
            foreach (string json in new[] { "private-malformed-marker", "{}", "null", "{\"SchemaVersion\":2,\"Profiles\":[]}",
                "{\"SchemaVersion\":1,\"Profiles\":null}", "{\"SchemaVersion\":1,\"Profiles\":[],\"Typo\":true}" })
            {
                var error = TestAssert.Throws<FormatException>(() => DoseRateEstimationProfileCatalog.Parse(json));
                TestAssert.False(error.ToString().Contains("private-malformed-marker"));
                TestAssert.True(error.InnerException == null);
            }
            var mutations = new Action<DoseRateEstimationProfile>[]
            {
                p => p.Id = null, p => p.Id = " ", p => p.Assumption = null,
                p => p.MaxGantrySpeedDegreesPerSecond = 0, p => p.MaxGantrySpeedDegreesPerSecond = -1,
                p => p.MaxGantrySpeedDegreesPerSecond = double.NaN, p => p.MaxGantrySpeedDegreesPerSecond = double.PositiveInfinity,
                p => p.MachineModelNames = null,
                p => { p.MachineModelNames = null; p.MlcModels = new List<string> { "Synthetic MLC" }; },
                p => p.MachineIds = new List<string> { " " },
                p => p.MachineModels = new List<string> { null },
                p => p.MachineModelNames = new List<string> { "Synthetic model " }
            };
            foreach (var mutate in mutations)
            {
                var profile = Profile(); mutate(profile); var invalid = Catalog(profile);
                TestAssert.True(invalid.Validate().Count > 0);
                TestAssert.Throws<FormatException>(() => DoseRateEstimationProfileCatalog.Parse(JsonConvert.SerializeObject(invalid)));
                var beam = Beam(); SeedStale(beam); DoseRateEstimator.Apply(beam, profile); Unavailable(beam);
            }
            var duplicate = Profile(); duplicate.Id = "SYNTHETIC-PROFILE";
            TestAssert.True(Catalog(Profile(), duplicate).Validate().Count > 0);
            TestAssert.True(Catalog((DoseRateEstimationProfile)null).Validate().Count > 0);
        }

        public static void OptionalFileSourceIsBoundedAndSanitized()
        {
            TestAssert.True(DoseRateProfileFileSource.LoadAsync(" ", CancellationToken.None).GetAwaiter().GetResult() == null);
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.Cancel();
                TestAssert.Throws<OperationCanceledException>(() => DoseRateProfileFileSource.LoadAsync(null, cancellation.Token).GetAwaiter().GetResult());
            }
            string directory = Path.Combine(Path.GetTempPath(), "ClearPlan-dose-rate-synthetic-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string path = Path.Combine(directory, "synthetic-profile.json");
                var catalog = Catalog(Profile()); catalog.Profiles[0].Assumption = "Synthetic UTF-8 assumption: Pr\u00fcfung.";
                File.WriteAllText(path, JsonConvert.SerializeObject(catalog), new UTF8Encoding(true));
                var loaded = DoseRateProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult();
                TestAssert.Equal(catalog.Profiles[0].Assumption, loaded.Profiles[0].Assumption);
                TestAssert.Equal(0, loaded.Validate().Count);
                foreach (string invalid in new[] { "private-malformed-marker", new string('x', 65537) })
                {
                    File.WriteAllText(path, invalid, new UTF8Encoding(false));
                    CheckFileError(path, directory);
                }
                File.WriteAllBytes(path, new byte[] { 0xc3, 0x28 }); // Invalid UTF-8 must not be replaced silently.
                CheckFileError(path, directory);
                File.Delete(path);
                CheckFileError(path, directory);
            }
            finally { Directory.Delete(directory, true); }
        }

        public static void ExtremePrecisionFailsClosed()
        {
            var beam = Beam(new[] { 0d, 1d }, new[] { 0d, 1d });
            beam.MetersetMu = 1e-100;
            DoseRateEstimator.Apply(beam, Profile(1e-308));
            Unavailable(beam); // Positive interval MU must not be rounded into a false zero-rate curve.
        }

        public static void TinyIntervalsRetainDirection()
        {
            var beam = Beam(new[] { 0d, 1e-16 }, new[] { 0d, 1d }); beam.MetersetMu = 1e-20;
            DoseRateEstimator.Apply(beam, Profile());
            TestAssert.Equal("Estimated", beam.DoseRateEstimateStatus);
            Near(0.036, beam.ControlPoints[1].EstimatedDoseRateMuPerMin);
            beam.GantryDirection = "CounterClockwise";
            DoseRateEstimator.Apply(beam, Profile()); Unavailable(beam);
        }

        private static void CheckFileError(string path, string privateDirectory)
        {
            var error = TestAssert.Throws<IOException>(() => DoseRateProfileFileSource.LoadAsync(path, CancellationToken.None).GetAwaiter().GetResult());
            TestAssert.False(error.ToString().Contains(privateDirectory) || error.ToString().Contains("private-malformed-marker"));
            TestAssert.True(error.InnerException == null);
        }

        private static ReviewBeamAnalysis Beam(double[] angles = null, double[] weights = null)
        {
            angles = angles ?? new[] { 0d, 12d, 24d }; weights = weights ?? new[] { 0d, 0.1d, 1d };
            return new ReviewBeamAnalysis
            {
                MachineId = "Synthetic-ID", MachineModelName = "Synthetic model", MachineModel = "Synthetic enum",
                MlcModel = "Synthetic MLC", Technique = "VMAT", GantryDirection = "Clockwise", MetersetMu = 10,
                NominalDoseRateMuPerMin = 540,
                ControlPoints = angles.Select((angle, index) => new ReviewControlPointSample
                { Index = index, GantryAngleDegrees = angle, CumulativeMetersetWeight = weights[index] }).ToList()
            };
        }
        private static DoseRateEstimationProfile Profile(double speed = 6)
        {
            return new DoseRateEstimationProfile { Id = "synthetic-profile", MachineModelNames = new List<string> { "Synthetic model" },
                MaxGantrySpeedDegreesPerSecond = speed, Assumption = "Synthetic test assumption only." };
        }
        private static DoseRateEstimationProfileCatalog Catalog(params DoseRateEstimationProfile[] profiles)
        { return new DoseRateEstimationProfileCatalog { SchemaVersion = 1, Profiles = profiles.ToList() }; }
        private static void SeedStale(ReviewBeamAnalysis beam)
        {
            beam.DoseRateEstimateStatus = "Estimated"; beam.DoseRateEstimateProfile = "stale";
            beam.EstimatedBeamDurationSeconds = 10; beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond = 6;
            foreach (var cp in beam.ControlPoints)
            { cp.EstimatedDoseRateMuPerMin = 500; cp.EstimatedSegmentDurationSeconds = 5; cp.PlannedDoseRateMuPerMin = 123; }
        }
        private static void Unavailable(ReviewBeamAnalysis beam)
        {
            TestAssert.Equal("Unavailable", beam.DoseRateEstimateStatus);
            TestAssert.True(!string.IsNullOrWhiteSpace(beam.DoseRateEstimateReason));
            TestAssert.False(beam.EstimatedBeamDurationSeconds.HasValue);
            TestAssert.False(beam.DoseRateEstimateMaxGantrySpeedDegreesPerSecond.HasValue);
            TestAssert.True(string.IsNullOrEmpty(beam.DoseRateEstimateProfile));
            if (beam.ControlPoints != null)
                foreach (var cp in beam.ControlPoints.Where(p => p != null))
                { TestAssert.False(cp.EstimatedDoseRateMuPerMin.HasValue); TestAssert.False(cp.EstimatedSegmentDurationSeconds.HasValue); }
        }
        private static void Near(double expected, double? actual)
        {
            TestAssert.True(actual.HasValue && !double.IsNaN(actual.Value) && !double.IsInfinity(actual.Value)
                && Math.Abs(expected - actual.Value) <= 1e-9 * Math.Max(1, Math.Abs(expected)),
                "Expected finite " + expected + ", found " + actual + ".");
        }
    }
}
