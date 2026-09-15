using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClearPlan.Core.Collision;
using ClearPlan.Core.PlanAnalysis;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class CollisionTests
    {
        public static void CoreContractExists()
        {
            TestAssert.NotNull(typeof(CollisionEvaluator),
                "A detached read-only collision evaluator is required.");
        }

        public static void SphereTriangleInteriorEdgesAndVertices()
        {
            var scene = Scene();
            Near(3, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
            SetCenter(scene, new BeamPoint3D(3, 3, 0));
            Near(Math.Sqrt(8) - 1, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
            SetCenter(scene, new BeamPoint3D(3, 3, 3));
            Near(Math.Sqrt(12) - 1, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
        }

        public static void SphereContainmentTangencyAndMargin()
        {
            var scene = Scene();
            scene.Poses[0].Isocenter = new BeamPoint3D(-5, 0, 0);
            scene.Poses[0].Source = new BeamPoint3D(1000, 0, 0);
            scene.Profile.HeadRadiusMm = 0.25;
            var result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            Near(-1.25, External(result).ClearanceMm);
            TestAssert.Equal("overlap", result.Status);
            scene = Scene(); SetCenter(scene, new BeamPoint3D(2, 0, 0));
            result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            Near(0, External(result).ClearanceMm); TestAssert.Equal("overlap", result.Status);
            SetCenter(scene, new BeamPoint3D(4, 0, 0));
            result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            Near(2, External(result).ClearanceMm); TestAssert.Equal("near", result.Status);
        }

        public static void SphereAnalyticBoxSweepAndTranslation()
        {
            var random = new Random(1234);
            for (int i = 0; i < 120; i++)
            {
                var center = new BeamPoint3D(random.NextDouble()*8-4, random.NextDouble()*8-4, random.NextDouble()*8-4);
                var scene = Scene(); SetCenter(scene, center);
                double dx = Math.Max(0,Math.Abs(center.X)-1), dy = Math.Max(0,Math.Abs(center.Y)-1), dz = Math.Max(0,Math.Abs(center.Z)-1);
                double expected = dx+dy+dz > 0 ? Math.Sqrt(dx*dx+dy*dy+dz*dz)-1
                    : -Math.Min(1-Math.Abs(center.X),Math.Min(1-Math.Abs(center.Y),1-Math.Abs(center.Z)))-1;
                Near(expected, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
                foreach (var surface in scene.Surfaces) foreach (var vertex in surface.Mesh.Vertices)
                { vertex.X += 1000; vertex.Y -= 220; vertex.Z += 350; }
                scene.Poses[0].Isocenter = new BeamPoint3D(1000,-220,350);
                scene.Poses[0].Source = new BeamPoint3D(center.X+1000,center.Y-220,center.Z+350);
                Near(expected, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
            }
        }

        public static void NumericallyCollapsedTrianglesCannotYieldClear()
        {
            var scene = Scene();
            // This nonzero-volume tetrahedron is far too ill-conditioned for stable double triangle distances.
            scene.Surfaces[0].Mesh = new TargetMeshGeometry {
                Vertices = new List<BeamPoint3D> { new BeamPoint3D(0,0,0),new BeamPoint3D(100000,0,0),
                    new BeamPoint3D(100000,0.0000001,0),new BeamPoint3D(0,0,1) },
                TriangleIndices = new[] { 0,2,1,0,1,3,0,3,2,1,2,3 } };
            Unavailable(scene);
        }

        public static void ContainmentEdgeRaysAndHollowComponents()
        {
            var scene = Scene();
            scene.Poses[0].Isocenter = new BeamPoint3D(-5,0,0);
            scene.Poses[0].Source = new BeamPoint3D(1000,0,0);
            // The first containment direction hits the (1, .371, .529) vertex exactly.
            scene.Surfaces[0].Mesh = Box(-1,-0.371,-0.529,1,0.371,0.529);
            Near(-1.371,External(CollisionEvaluator.Evaluate(scene,CancellationToken.None)).ClearanceMm);
            var outer = Box(-3,-3,-3,3,3,3); var inner = Box(-1,-1,-1,1,1,1);
            int offset = outer.Vertices.Count; outer.Vertices.AddRange(inner.Vertices);
            // A closed inner boundary encloses a cavity: parity is even in that cavity.
            outer.TriangleIndices = outer.TriangleIndices.Concat(inner.TriangleIndices.Reverse().Select(i => i+offset)).ToArray();
            scene.Surfaces[0].Mesh = outer; scene.Profile.HeadRadiusMm = 0.25;
            Near(0.75,External(CollisionEvaluator.Evaluate(scene,CancellationToken.None)).ClearanceMm);
            // Global winding reversal must not alter the physical signed distance.
            outer.TriangleIndices = outer.TriangleIndices.Reverse().ToArray();
            Near(0.75,External(CollisionEvaluator.Evaluate(scene,CancellationToken.None)).ClearanceMm);
        }

        public static void RingTangencyTranslationAndPartialCoverage()
        {
            var scene = Scene("RingBore");
            scene.Surfaces[0].Mesh = Box(-3,-4,-2,3,4,2);
            scene.Surfaces[1].Mesh = Box(-1,-1,-2,1,1,2);
            scene.Profile.BoreRadiusMm = 5;
            var result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            Near(0, External(result).ClearanceMm); TestAssert.Equal("overlap", result.Status);
            foreach (var surface in scene.Surfaces) foreach (var vertex in surface.Mesh.Vertices)
            { vertex.X += 300; vertex.Y += 200; vertex.Z -= 100; }
            scene.Poses[0].Isocenter = new BeamPoint3D(300,200,-100);
            scene.Poses[0].Source = new BeamPoint3D(1300,200,-100);
            Near(0, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
            foreach (var vertex in scene.Surfaces[1].Mesh.Vertices) vertex.Z += 100;
            result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            TestAssert.Equal("unavailable", result.Status);
            TestAssert.True(result.Findings.Any(f => f.Status == "overlap"), "Partial coverage must preserve individual overlap warnings.");
        }

        public static void RingClipsTrianglesToFiniteAxialSlab()
        {
            var scene = Scene("RingBore");
            scene.Surfaces[0].Mesh = Box(-3, -4, -2, 3, 4, 2);
            scene.Surfaces[1].Mesh = Box(-2, -2, -2, 2, 2, 2);
            var result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            Near(5, External(result).ClearanceMm); TestAssert.Equal("sampled-clear", result.Status);
            // Only the triangle/slab intersection counts: the distant vertex at z=10 must be clipped.
            scene.Profile.BoreRadiusMm = 20;
            scene.Surfaces[0].Mesh = new TargetMeshGeometry {
                Vertices = new List<BeamPoint3D> { new BeamPoint3D(100,0,10), new BeamPoint3D(2,0,0),
                    new BeamPoint3D(0,2,0), new BeamPoint3D(0,0,0) },
                TriangleIndices = new[] { 0,2,1, 0,1,3, 0,3,2, 1,2,3 } };
            Near(20 - 11.8, External(CollisionEvaluator.Evaluate(scene, CancellationToken.None)).ClearanceMm);
            foreach (var vertex in scene.Surfaces[0].Mesh.Vertices) vertex.Z += 50;
            result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            TestAssert.Equal("unavailable", result.Status);
            TestAssert.False(External(result).ClearanceMm.HasValue, "No axial coverage must not appear as infinite clearance.");
        }

        public static void MissingUncommissionedAndInvalidInputsFailClosed()
        {
            TestAssert.Equal("unavailable", CollisionEvaluator.Evaluate(null, CancellationToken.None).Status);
            var scene = Scene(); scene.SupportedCoordinates = false; Unavailable(scene);
            scene = Scene(); scene.Profile = null; Unavailable(scene);
            scene = Scene(); scene.Profile.Commissioned = false; Unavailable(scene);
            scene = Scene(); scene.Profile.SyntheticOnly = true; scene.Synthetic = false; Unavailable(scene);
            scene = Scene(); scene.Profile.Evidence = ""; Unavailable(scene);
            scene = Scene(); scene.Profile.HeadRadiusMm = double.NaN; Unavailable(scene);
            scene = Scene(); scene.Profile.SafetyMarginMm = -1; Unavailable(scene);
            scene = Scene(); scene.Surfaces.RemoveAt(1); Unavailable(scene);
            scene = Scene(); scene.Surfaces.RemoveAt(0); Unavailable(scene);
            scene = Scene(); scene.Poses.Clear(); Unavailable(scene);
            scene = Scene(); scene.Poses[0].Source = scene.Poses[0].Isocenter; Unavailable(scene);
            scene = Scene(); scene.Poses[0].GantryDegrees = double.NaN; Unavailable(scene);
            scene = Scene(); scene.Profile.SyntheticOnly = true; scene.Profile.Commissioned = false;
            TestAssert.Equal("sampled-clear", CollisionEvaluator.Evaluate(scene, CancellationToken.None).Status);
        }

        public static void MeshTopologyNumericsAndWorkloadFailClosed()
        {
            TestAssert.Equal("sampled-clear", CollisionEvaluator.Evaluate(Scene(), CancellationToken.None).Status);
            var scene = Scene(); scene.Surfaces[0].Mesh.Vertices.Clear(); Unavailable(scene);
            scene = Scene(); scene.Surfaces[0].Mesh.TriangleIndices[0] = 1000; Unavailable(scene);
            scene = Scene(); scene.Surfaces[0].Mesh.TriangleIndices = new[] { 0, 1, 2 }; Unavailable(scene);
            scene = Scene(); scene.Surfaces[0].Mesh.Vertices[0].X = double.PositiveInfinity; Unavailable(scene);
            scene = Scene(); scene.Surfaces[0].Mesh.Vertices[0] = scene.Surfaces[0].Mesh.Vertices[1]; Unavailable(scene);
            scene = Scene(); scene.Surfaces[0].Mesh.TriangleIndices = scene.Surfaces[0].Mesh.TriangleIndices.Concat(new[] { 0, 2, 1 }).ToArray(); Unavailable(scene);
            scene = Scene(); scene.Poses = Enumerable.Repeat(scene.Poses[0], 10001).ToList(); Unavailable(scene);
        }

        public static void ProfilesAreBoundedExactAndEvidenceGated()
        {
            var profile = Profile("CArmSphere");
            var catalog = CollisionProfileCatalog.Parse(JsonConvert.SerializeObject(new { SchemaVersion = 1, Profiles = new[] { profile } }));
            TestAssert.NotNull(catalog, "Valid explicit profile catalog must be parsed.");
            TestAssert.NotNull(catalog.Select("test-machine"));
            TestAssert.Equal(null, catalog.Select("MLC120"));
            TestAssert.Equal(null, catalog.Select("TEST"));
            TestAssert.Equal(null, catalog.Select(" TEST-MACHINE"));
            TestAssert.Equal(null, CollisionProfileCatalog.Parse("{\"SchemaVersion\":1,\"Profiles\":[]}").Select("TEST-MACHINE"));
            TestAssert.Throws<ArgumentException>(() => CollisionProfileCatalog.Parse("{\"SchemaVersion\":2,\"schemaVersion\":1,\"Profiles\":[]}"));
            string conflictingProfile = JsonConvert.SerializeObject(new { SchemaVersion = 1, Profiles = new[] { profile } })
                .Replace("\"Commissioned\":true", "\"Commissioned\":false,\"commissioned\":true");
            TestAssert.Throws<ArgumentException>(() => CollisionProfileCatalog.Parse(conflictingProfile));
            foreach (string json in new[] { "null", "{}", "{\"SchemaVersion\":2,\"Profiles\":[]}",
                "{\"SchemaVersion\":1,\"Profiles\":[],\"Unknown\":true}", "{\"SchemaVersion\":1,\"SchemaVersion\":2,\"Profiles\":[]}",
                new string(' ', 1024*1024+1) })
                TestAssert.Throws<ArgumentException>(() => CollisionProfileCatalog.Parse(json));
            TestAssert.Throws<ArgumentException>(() => CollisionProfileCatalog.Parse(JsonConvert.SerializeObject(new { SchemaVersion = 1, Profiles = new[] { profile, profile } })));
            profile.Evidence = null;
            TestAssert.Throws<ArgumentException>(() => CollisionProfileCatalog.Validate(profile, false));
        }

        public static void CancellationAndSyntheticLimitations()
        {
            TestAssert.Throws<OperationCanceledException>(() => CollisionEvaluator.Evaluate(Scene(), new CancellationToken(true)));
            var scene = SyntheticCollisionFactory.Create("SYNTHETIC");
            TestAssert.True(scene.Synthetic && scene.Profile.SyntheticOnly && !scene.Profile.Commissioned);
            var result = CollisionEvaluator.Evaluate(scene, CancellationToken.None);
            TestAssert.True(result.Findings.Any(f => f.Status == "overlap"));
            TestAssert.True(result.Findings.Any(f => f.Status == "near"));
            TestAssert.True(result.Findings.Any(f => f.Status == "sampled-clear"));
            foreach (string limitation in new[] { "sampled", "interpolation", "CT truncation", "accessories", "setup", "insertion" })
                TestAssert.True(result.Summary.Contains(limitation), "Research preview must state " + limitation + " limitation.");
            scene.Synthetic = false; Unavailable(scene);
            var ring = SyntheticCollisionFactory.Create("SYNTHETIC", "RingBore");
            TestAssert.True(ring.Synthetic && ring.Profile.Kind == "RingBore");
            TestAssert.True(CollisionEvaluator.Evaluate(ring, CancellationToken.None).Findings.Count > 0);
            TestAssert.Throws<ArgumentException>(() => SyntheticCollisionFactory.Create("SYNTHETIC", "Unknown"));
        }

        public static void WorkloadLimitAndInFlightCancellation()
        {
            var scene = SyntheticCollisionFactory.Create("SYNTHETIC");
            scene.Poses = Enumerable.Range(0,10000).Select(i => new CollisionPose { BeamId = "SYNTHETIC", ControlPointIndex = i,
                Isocenter = new BeamPoint3D(0,0,0), Source = new BeamPoint3D(1000,0,0) }).ToList();
            // Two body surfaces produce a workload above ten million triangle/pose pairs.
            scene.Surfaces.Add(new CollisionSurface { Label = "SYNTHETIC extra", Role = "External", Mesh = scene.Surfaces[0].Mesh });
            Unavailable(scene);
            double originalX = scene.Surfaces[0].Mesh.Vertices[0].X;
            scene.Surfaces[0].Mesh.Vertices[0].X = double.NaN;
            TestAssert.True(CollisionEvaluator.Evaluate(scene,CancellationToken.None).Summary.Contains("workload"),
                "Oversized work must be rejected before allocating/validating complete triangle topology.");
            scene.Surfaces[0].Mesh.Vertices[0].X = originalX;
            scene.Surfaces.RemoveAt(2);
            using (var cancellation = new CancellationTokenSource())
            {
                cancellation.CancelAfter(10);
                TestAssert.Throws<OperationCanceledException>(() => CollisionEvaluator.Evaluate(scene,cancellation.Token));
            }
        }

        private static CollisionScene Scene(string kind = "CArmSphere")
        {
            return new CollisionScene { PlanKey = "SYNTHETIC", Synthetic = true, SupportedCoordinates = true, Profile = Profile(kind),
                Surfaces = new List<CollisionSurface> {
                    new CollisionSurface { Label = "body", Role = "External", Mesh = Box(-1,-1,-1,1,1,1) },
                    new CollisionSurface { Label = "table", Role = "Support", Mesh = Box(-50,-2,-2,-40,2,2) } },
                Poses = new List<CollisionPose> { new CollisionPose { BeamId = "SYNTHETIC", Isocenter = new BeamPoint3D(0,0,0),
                    Source = new BeamPoint3D(1000,0,0), GantryDegrees = 90 } } };
        }
        private static CollisionProfile Profile(string kind)
        {
            return new CollisionProfile { MachineId = "TEST-MACHINE", Kind = kind, Revision = "test-1", Commissioned = true,
                CommissionedBy = "Synthetic unit test", Evidence = "Analytic test only", HeadCenterFromIsoMm = 5,
                HeadRadiusMm = 1, BoreRadiusMm = 10, BoreHalfLengthMm = 1, SafetyMarginMm = 2 };
        }
        private static void SetCenter(CollisionScene scene, BeamPoint3D center)
        {
            scene.Poses[0].Source = center;
            scene.Profile.HeadCenterFromIsoMm = Math.Sqrt(center.X*center.X + center.Y*center.Y + center.Z*center.Z);
        }
        private static CollisionFinding External(CollisionResult result)
        {
            var finding = result.Findings.FirstOrDefault(f => f.SurfaceLabel == "body");
            TestAssert.NotNull(finding, "Expected evaluated External triangle geometry."); return finding;
        }
        private static void Near(double expected, double? actual)
        { TestAssert.True(actual.HasValue && Math.Abs(expected-actual.Value) < 1e-6, "Expected clearance " + expected + "; found " + actual); }
        private static void Unavailable(CollisionScene scene)
        { TestAssert.Equal("unavailable", CollisionEvaluator.Evaluate(scene, CancellationToken.None).Status); }
        private static TargetMeshGeometry Box(double x0, double y0, double z0, double x1, double y1, double z1)
        {
            return new TargetMeshGeometry { Vertices = new List<BeamPoint3D> {
                new BeamPoint3D(x0,y0,z0),new BeamPoint3D(x1,y0,z0),new BeamPoint3D(x1,y1,z0),new BeamPoint3D(x0,y1,z0),
                new BeamPoint3D(x0,y0,z1),new BeamPoint3D(x1,y0,z1),new BeamPoint3D(x1,y1,z1),new BeamPoint3D(x0,y1,z1) },
                TriangleIndices = new[] { 0,2,1,0,3,2,4,5,6,4,6,7,0,1,5,0,5,4,1,2,6,1,6,5,2,3,7,2,7,6,3,0,4,3,4,7 } };
        }
#if COLLISION_STANDALONE
        private static int Main()
        {
            int failures = 0, count = 0;
            foreach (var method in typeof(CollisionTests).GetMethods(System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.Static))
            {
                try { method.Invoke(null, null); Console.WriteLine("PASS " + method.Name); }
                catch (Exception error) { failures++; Console.WriteLine("FAIL " + method.Name + ": " + (error.InnerException ?? error).Message); }
                count++;
            }
            Console.WriteLine(count + " collision groups, " + failures + " failed."); return failures == 0 ? 0 : 1;
        }
#endif
    }
}
