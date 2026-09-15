using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using ClearPlan.Core.Collision;
using ClearPlan.Core.PlanAnalysis;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class SourceCollisionScreeningTests
    {
        public static void SeparateContractExists()
        {
            var assembly = typeof(CollisionScene).Assembly;
            foreach (var name in new[] { "SourceCollisionModelCatalog", "SourceCollisionModel", "SourceCollisionScreening", "SourceCollisionScreeningRequest", "SourceCollisionScreeningResult" })
                TestAssert.NotNull(assembly.GetType("ClearPlan.Core.Collision." + name), "Separate source-screening API is missing: " + name);
        }

        public static void CatalogPreservesUnitsProvenanceAndExactBinding()
        {
            var catalog = Catalog();
            var model = catalog.Select("native-tb");
            TestAssert.Equal(420.0, model.HeadObstacles[0].FrontFaceFromIsoMm);
            TestAssert.Equal(370.0, model.HeadObstacles[0].RadiusMm);
            TestAssert.Equal("source-recorded-measurement", model.EvidenceLevel);
            TestAssert.Equal("2026-07-30", model.MeasurementDate);
            TestAssert.True(catalog.Select("native-tb-extra") == null);
            TestAssert.True(catalog.Select(" native-tb") == null);
            TestAssert.True(catalog.Select("Ethos") == null, "No inferred Ethos mapping.");
            TestAssert.True(catalog.Select(null) == null);
            TestAssert.NotNull(catalog.Select("NATIVE-TB"));
        }

        public static void CatalogRejectsAmbiguousMalformedAndCommissioningClaims()
        {
            Action<string> reject = json => TestAssert.Throws<ArgumentException>(() => SourceCollisionModelCatalog.Parse(json));
            var raw = RawCatalog();
            raw.Units = "cm"; reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.Models[0].HeadObstacles[0].RadiusMm = double.NaN; reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.Models[0].SourceSha256 = ""; reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.Models[0].MeasurementDate = null; reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.MachineBindings.Add(raw.MachineBindings[0]); reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.MachineBindings[0].ModelId = "missing"; reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.Models.Add(raw.Models[0]); reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.Models[0].ReachRadiusMm = 0; reject(JsonConvert.SerializeObject(raw));
            raw = RawCatalog(); raw.Models[1].BoreWarningRadiusMm = 501; reject(JsonConvert.SerializeObject(raw));
            reject("{}"); reject("{"); reject("{\"SchemaVersion\":1,\"schemaversion\":1}");
            reject(JsonConvert.SerializeObject(RawCatalog()).Replace("\"SchemaVersion\":1", "\"SchemaVersion\":1,\"Commissioned\":true"));
            reject(new string(' ', 1024 * 1024 + 1));
        }

        public static void TrueBeamUsesMillimetresAndIncludesTangency()
        {
            CheckPoint(new BeamPoint3D(0, -350, 0), 0, 0);
            CheckPoint(new BeamPoint3D(0, -390, 0), 0, 1);
            CheckPoint(new BeamPoint3D(0, -400, 0), 1, 1); // pin-front tangency
            CheckPoint(new BeamPoint3D(200, -400, 0), 1, 1); // pin-rim tangency
            CheckPoint(new BeamPoint3D(201, -400, 0), 0, 1);
            CheckPoint(new BeamPoint3D(370, -420, 0), 1, 1); // cover-rim tangency
            CheckPoint(new BeamPoint3D(371, -420, 0), 0, 1);
            CheckPoint(new BeamPoint3D(0, 500, 0), 0, 0); // behind iso, not source
            var request = Request(new BeamPoint3D(0, -420, 0));
            var result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal("screening-hit", result.Status);
            TestAssert.False(result.ClinicalClearanceSupported);
            TestAssert.True(result.Summary.Contains("surface-point") && result.Summary.Contains("captured poses"));
            TestAssert.True(result.Summary.Contains("not clinical clearance"));
        }

        public static void TrueBeamNativeDirectionAndTranslationsDeterminePose()
        {
            var request = Request(new BeamPoint3D(500, 0, 0));
            request.Poses[0].Source = new BeamPoint3D(1000, 0, 0);
            var result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal(1, result.Findings[0].ModelHitPointCount);
            request.Poses[0].GantryDegrees = 271; // metadata does not override native direction
            TestAssert.Equal(1, SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Findings[0].ModelHitPointCount);
            request.Poses[0].Isocenter = new BeamPoint3D(10, 20, 30);
            request.Poses[0].Source = new BeamPoint3D(1010, 20, 30);
            request.Surfaces[0].Points[0] = new BeamPoint3D(510, 20, 30);
            TestAssert.Equal(1, SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Findings[0].ModelHitPointCount);
            request = Request(new BeamPoint3D(0, 0, 500));
            request.Poses[0].Source = new BeamPoint3D(0, 0, 1000);
            TestAssert.Equal(1, SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Findings[0].ModelHitPointCount);
            request = Request(new BeamPoint3D(500 / Math.Sqrt(2), -500 / Math.Sqrt(2), 0));
            request.Poses[0].Source = new BeamPoint3D(1000, -1000, 0);
            TestAssert.Equal(1, SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Findings[0].ModelHitPointCount);
            request = Request(new BeamPoint3D(500, 0, 0));
            request.Poses.Add(new CollisionPose { BeamId = "B", ControlPointIndex = 1, GantryDegrees = 90,
                Isocenter = new BeamPoint3D(0,0,0), Source = new BeamPoint3D(1000,0,0) });
            result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal(0, result.Findings[0].ModelHitPointCount);
            TestAssert.Equal(1, result.Findings[1].ModelHitPointCount);
            TestAssert.Equal(1, result.Findings[1].ControlPointIndex);
        }

        public static void ReachAndSampleScopeCannotClaimClearance()
        {
            CheckPoint(new BeamPoint3D(0, -700, 0), 1, 1);
            var request = Request(new BeamPoint3D(0, -701, 0));
            var result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal("unavailable", result.Status);
            TestAssert.Equal(1, result.Findings[0].ExcludedPointCount);
            request.Surfaces[0].Points.Add(new BeamPoint3D(0, -100, 0));
            result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal("screening-no-hit", result.Status);
            TestAssert.Equal(1, result.Findings[0].ExcludedPointCount);
            TestAssert.Equal(1, result.Findings[0].ScreenedPointCount);
            TestAssert.False(result.ClinicalClearanceSupported);
            TestAssert.Equal("source-recorded-measurement", result.EvidenceLevel);
            TestAssert.Equal(new string('a', 64), result.SourceSha256);
        }

        public static void HalcyonUsesSourceRadiiLeadingLimitsAndExternalOnly()
        {
            CheckRing(new BeamPoint3D(479, 0, 0), "external", 0, 0);
            CheckRing(new BeamPoint3D(480, 0, 0), "external", 0, 1);
            CheckRing(new BeamPoint3D(500, 0, 0), "external", 1, 1);
            CheckRing(new BeamPoint3D(0, 0, 850), "external", 0, 1);
            CheckRing(new BeamPoint3D(0, 0, 900), "external", 1, 1);
            CheckRing(new BeamPoint3D(0, 0, -1000), "external", 0, 0);
            CheckRing(new BeamPoint3D(0, 0, 1000), "support", 0, 0);
            CheckRing(new BeamPoint3D(501, 0, 0), "support", 1, 1);
            var request = Request(new BeamPoint3D(510, 20, 30)); request.MachineId = "native-hal";
            request.Poses[0].Isocenter = new BeamPoint3D(10, 20, 30);
            request.Poses[0].Source = new BeamPoint3D(10, -980, 30);
            var result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal(1, result.Findings[0].ModelHitPointCount);
            TestAssert.Equal("source-constants", result.EvidenceLevel);
        }

        public static void MissingInvalidNonNativeAndUnsupportedInputsAreUnavailable()
        {
            Unavailable(null);
            var request = Request(); request.MachineId = null; Unavailable(request);
            request = Request(); request.MachineId = "Ethos"; Unavailable(request);
            request = Request(); request.NativeSourceCoordinates = false; Unavailable(request);
            request = Request(); request.CoordinateSystem = "DICOM-LPS-cm"; Unavailable(request);
            request = Request(); request.CoordinateSystem = null; Unavailable(request);
            request = Request(); request.PatientPosition = ""; Unavailable(request);
            request = Request(); request.PatientPosition = "FFS"; request.MachineId = "native-hal"; Unavailable(request);
            request = Request(); request.Surfaces = null; Unavailable(request);
            request = Request(); request.Surfaces[0].Points.Clear(); Unavailable(request);
            request = Request(); request.Surfaces[0].Role = "unknown"; Unavailable(request);
            request = Request(); request.Surfaces[0].Points[0] = null; Unavailable(request);
            request = Request(); request.Surfaces[0].Points[0].X = double.PositiveInfinity; Unavailable(request);
            request = Request(); request.Surfaces[0].Points[0].Y = double.NaN; Unavailable(request);
            request = Request(); request.Poses = null; Unavailable(request);
            request = Request(); request.Poses[0].Isocenter = null; Unavailable(request);
            request = Request(); request.Poses[0].Source = null; Unavailable(request);
            request = Request(); request.Poses[0].Source = request.Poses[0].Isocenter; Unavailable(request);
            request = Request(); request.Poses[0].GantryDegrees = double.NaN; Unavailable(request);
            request = Request(); request.Poses[0].ControlPointIndex = -1; Unavailable(request);
            request = Request(); request.Surfaces[0].Points = Enumerable.Repeat(new BeamPoint3D(0,0,0), 200001).ToList(); Unavailable(request);
            TestAssert.Equal("unavailable", SourceCollisionScreening.Evaluate(Request(), null, CancellationToken.None).Status);
            var catalog = Catalog(); catalog.Models[0].HeadObstacles[0].RadiusMm = -1;
            TestAssert.Equal("unavailable", SourceCollisionScreening.Evaluate(Request(), catalog, CancellationToken.None).Status);
        }

        public static void CancellationNeverReturnsScreeningSuccess()
        {
            using (var cancel = new CancellationTokenSource())
            {
                cancel.Cancel();
                var result = SourceCollisionScreening.Evaluate(Request(), Catalog(), cancel.Token);
                TestAssert.Equal("cancelled", result.Status);
                TestAssert.Equal(0, result.Findings.Count);
                TestAssert.False(result.ClinicalClearanceSupported);
            }
        }

        public static void DisabledCatalogAndPublishedSourceAreExplicit()
        {
            var empty = SourceCollisionModelCatalog.Parse("{\"SchemaVersion\":1,\"Units\":\"mm\",\"Models\":[],\"MachineBindings\":[]}");
            TestAssert.True(empty.Select("native-tb") == null);
            TestAssert.Equal("unavailable", SourceCollisionScreening.Evaluate(Request(), empty, CancellationToken.None).Status);
            var dir = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory);
            while (dir != null && !Directory.Exists(Path.Combine(dir.FullName, "ClearPlan.Core"))) dir = dir.Parent;
            TestAssert.NotNull(dir);
            var path = Path.Combine(dir.FullName, "ClearPlan.Script", "Distribution", "MachineGeometry", "SourceCollisionModels.example.json");
            TestAssert.True(File.Exists(path), "Empty public source-model catalog is missing.");
            var models = SourceCollisionModelCatalog.Parse(File.ReadAllText(path));
            TestAssert.Equal(0, models.MachineBindings.Count, "Machine bindings require explicit local configuration.");
            TestAssert.Equal(0, models.Models.Count, "The public catalog must not supply private or uncommissioned machine dimensions.");
            TestAssert.True(models.Select("example-device") == null);
            TestAssert.True(models.Select("Ethos") == null);
        }

        public static void NativeWorkloadIsExplicitBoundedAndCancellableInFlight()
        {
            // 500 captured poses x 200k points is supported for screening, not silently clipped.
            var request = Request();
            request.Surfaces[0].Points = Enumerable.Repeat(new BeamPoint3D(0,-100,0), 200000).ToList();
            request.Poses = Enumerable.Range(0, 500).Select(i => Pose(i)).ToList();
            using (var cancel = new CancellationTokenSource())
            using (var timer = new Timer(_ => cancel.Cancel(), null, 20, Timeout.Infinite))
            {
                var result = SourceCollisionScreening.Evaluate(request, Catalog(), cancel.Token);
                TestAssert.Equal("cancelled", result.Status, "A realistic bounded workload should run until cancellation, not be rejected by a tiny cap.");
                TestAssert.Equal(0, result.Findings.Count);
            }
            request.Poses.Add(Pose(500));
            Unavailable(request); // 100.2m point-pose pairs exceed the explicit maximum.
            request = Request(); request.Poses = Enumerable.Repeat(request.Poses[0], 10001).ToList(); Unavailable(request);
        }

        public static void RotatedTangencyAndResultSizeAreBounded()
        {
            var request = Request(new BeamPoint3D(400 / Math.Sqrt(2), -400 / Math.Sqrt(2), 0));
            request.Poses[0].Source = new BeamPoint3D(1000, -1000, 0);
            TestAssert.Equal(1, SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Findings.Single().ModelHitPointCount,
                "A rotated tangent must not disappear through floating point roundoff.");
            request = Request(new BeamPoint3D(700 / Math.Sqrt(2), -700 / Math.Sqrt(2), 0));
            request.Poses[0].Source = new BeamPoint3D(1000, -1000, 0);
            TestAssert.Equal(1, SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Findings.Single().ScreenedPointCount);
            request = Request(); request.Surfaces = Enumerable.Range(0, 32).Select(i => new SourceCollisionSurface {
                Label = "surface " + i, Role = "external", Points = new List<BeamPoint3D> { new BeamPoint3D(0,-100,0) } }).ToList();
            request.Poses = Enumerable.Range(0, 313).Select(i => Pose(i)).ToList();
            Unavailable(request); // Output cardinality is bounded independently of point-work.
        }

        public static void DuplicateIdentitiesFailAndMissingSurfaceRolesAreDisclosed()
        {
            var request = Request(); request.Poses.Add(Pose(0)); Unavailable(request);
            request = Request(); var duplicate = Pose(0); duplicate.BeamId = "b"; request.Poses.Add(duplicate); Unavailable(request);
            request = Request(); request.Surfaces.Add(new SourceCollisionSurface { Label = "SYNTHETIC SURFACE", Role = "support",
                Points = new List<BeamPoint3D> { new BeamPoint3D(0,-100,0) } }); Unavailable(request);
            var result = SourceCollisionScreening.Evaluate(Request(), Catalog(), CancellationToken.None);
            TestAssert.Equal("screening-no-hit", result.Status);
            TestAssert.True(result.Summary.Contains("support missing"), "A missing support must be disclosed without claiming complete screening.");
            request = Request(); request.Surfaces[0].Role = "support";
            result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.True(result.Summary.Contains("external missing"));
            TestAssert.False(result.ClinicalClearanceSupported);
        }

        private static void CheckPoint(BeamPoint3D point, int hits, int margin)
        {
            var result = SourceCollisionScreening.Evaluate(Request(point), Catalog(), CancellationToken.None);
            TestAssert.Equal(hits, result.Findings.Single().ModelHitPointCount);
            TestAssert.Equal(margin, result.Findings.Single().MarginHitPointCount);
        }

        private static void CheckRing(BeamPoint3D point, string role, int hits, int margin)
        {
            var request = Request(point); request.MachineId = "native-hal"; request.Surfaces[0].Role = role;
            var result = SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None);
            TestAssert.Equal(hits, result.Findings.Single().ModelHitPointCount);
            TestAssert.Equal(margin, result.Findings.Single().MarginHitPointCount);
        }

        private static void Unavailable(SourceCollisionScreeningRequest request)
        { TestAssert.Equal("unavailable", SourceCollisionScreening.Evaluate(request, Catalog(), CancellationToken.None).Status); }

        private static SourceCollisionScreeningRequest Request(BeamPoint3D point = null)
        {
            return new SourceCollisionScreeningRequest {
                MachineId = "native-tb", CoordinateSystem = "DICOM-LPS-mm", PatientPosition = "HFS", NativeSourceCoordinates = true,
                Surfaces = new List<SourceCollisionSurface> { new SourceCollisionSurface { Label = "synthetic surface", Role = "external",
                    Points = new List<BeamPoint3D> { point ?? new BeamPoint3D(0, -100, 0) } } },
                Poses = new List<CollisionPose> { Pose(0) } };
        }

        private static CollisionPose Pose(int index)
        { return new CollisionPose { BeamId = "B", ControlPointIndex = index, GantryDegrees = 0,
            Isocenter = new BeamPoint3D(0,0,0), Source = new BeamPoint3D(0,-1000,0) }; }

        private static SourceCollisionModelCatalog Catalog()
        { return SourceCollisionModelCatalog.Parse(JsonConvert.SerializeObject(RawCatalog())); }

        private static SourceCollisionModelCatalog RawCatalog()
        {
            return new SourceCollisionModelCatalog { SchemaVersion = 1, Units = "mm",
                MachineBindings = new List<SourceCollisionMachineBinding> {
                    new SourceCollisionMachineBinding { MachineId = "native-tb", ModelId = "tb-source" },
                    new SourceCollisionMachineBinding { MachineId = "native-hal", ModelId = "hal-source" } },
                Models = new List<SourceCollisionModel> {
                    new SourceCollisionModel { ModelId = "tb-source", Kind = "TrueBeamHeadSource", SourcePath = "synthetic-source.py",
                        SourceSha256 = new string('a',64), SourceSnapshotDate = "2026-09-13", EvidenceLevel = "source-recorded-measurement",
                        MeasurementDate = "2026-07-30", Evidence = "Synthetic fixture reflecting source dimensions; no commissioning.",
                        ReachRadiusMm = 700, WarningMarginMm = 30, SourceAngularMarginDegrees = 5,
                        HeadObstacles = new List<SourceHeadObstacle> { new SourceHeadObstacle { FrontFaceFromIsoMm = 420, RadiusMm = 370 },
                            new SourceHeadObstacle { FrontFaceFromIsoMm = 400, RadiusMm = 200 } } },
                    new SourceCollisionModel { ModelId = "hal-source", Kind = "HalcyonBoreSource", SourcePath = "synthetic-source.py",
                        SourceSha256 = new string('a',64), SourceSnapshotDate = "2026-09-13", EvidenceLevel = "source-constants",
                        Evidence = "Synthetic fixture reflecting unverified source thresholds.", BoreWarningRadiusMm = 480,
                        BoreLimitRadiusMm = 500, LeadingWarningFromIsoMm = 850, LeadingLimitFromIsoMm = 900 } } };
        }
    }
}
