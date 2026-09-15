using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Threading;
using ClearPlan.Core.Collision;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Tests
{
    internal static class CollisionSweepTests
    {
        public static void ModelSeparationDoesNotChangeClinicalAssurance()
        {
            var property = typeof(CollisionSweepFrame).GetProperty("ModelBodyStatus");
            TestAssert.NotNull(property, "Captured model separation must be distinct from commissioning/CT assurance.");
            var scene = SourceScene(); scene.SupportedCoordinates = false; scene.NominalCoordinatesSupported = true;
            scene.CtCoverageKnown = true; scene.ExternalTruncatedAtCtBoundary = true;
            scene.Surfaces[0].Mesh = Box(50,-1,-1,51,1,1);
            var frame = CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("model-clear", property.GetValue(frame));
            TestAssert.Equal("uncertain",frame.Status,"The clinical/coverage gate must not be weakened.");
            // WPF can include unsupported degenerate primitives. If the distance
            // domain includes them, a positive bound still proves separation of
            // the supplied surface, not anatomical completeness.
            scene.Surfaces[0].Mesh.Vertices.Add(new BeamPoint3D(40,0,0));
            scene.Surfaces[0].Mesh.TriangleIndices=scene.Surfaces[0].Mesh.TriangleIndices.Concat(new[]{8,8,8}).ToArray();
            frame=CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("model-clear",property.GetValue(frame),"Positive bounds cover degenerate primitives too; the open-bore validity gate is not a head surface-distance prerequisite.");
            scene.Surfaces[0].Mesh.Vertices[8]=new BeamPoint3D(110,0,0);
            scene.Surfaces[0].Mesh.TriangleIndices=Enumerable.Range(0,120).SelectMany(i=>scene.Surfaces[0].Mesh.TriangleIndices).ToArray();
            frame=CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("model-hit",property.GetValue(frame),"Spatial pruning must retain a degenerate inside-point witness beside proper surface faces.");
            scene.Surfaces[0].Mesh=PointTriangle(40,0,0);
            frame=CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("uncertain",property.GetValue(frame),"A point alone is not a captured body surface.");
            scene.Surfaces[0].Mesh = Box(110,-1,-1,111,1,1);
            frame = CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("model-hit",property.GetValue(frame));
            TestAssert.Equal("uncertain",frame.Status);
            scene.NominalCoordinatesSupported = false;
            frame = CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("uncertain",property.GetValue(frame));
            scene.NominalCoordinatesSupported = true; scene.Surfaces[0].Mesh = null;
            frame = CollisionSweep.Evaluate(scene,CancellationToken.None).Frames.Single();
            TestAssert.Equal("uncertain",property.GetValue(frame));
            var arc = Scene(3,23); var sampled = CollisionSweep.Evaluate(arc,CancellationToken.None).Frames;
            var identity = typeof(CollisionSweepFrame).GetProperty("CapturedControlPointIndex");
            TestAssert.NotNull(identity);
            TestAssert.Equal(0,identity.GetValue(sampled[1]));
            TestAssert.Equal(1,identity.GetValue(sampled[2]));
            TestAssert.Equal(1,identity.GetValue(sampled[3]));
        }

        public static void DetachedContractExists()
        {
            var type = typeof(CollisionScene).Assembly.GetType("ClearPlan.Core.Collision.CollisionSweep");
            TestAssert.NotNull(type, "A bounded detached 10-degree sweep with per-role signed distances is required.");
            foreach (var name in new[] { "CtCoverageKnown", "ExternalTruncatedAtCtBoundary", "CtCoverageReason" })
                TestAssert.NotNull(typeof(CollisionScene).GetProperty(name), "Explicit CT coverage metadata is required: " + name);
        }

        public static void ArcOrderEndpointsAndStaticDeduplication()
        {
            AssertAngles(Scene(350, 10), 350, 0, 10);
            AssertAngles(Scene(10, 350), 10, 0, 350);
            AssertAngles(Scene(3, 23), 3, 10, 20, 23);
            AssertAngles(Scene(23, 3), 23, 20, 10, 3);
            AssertAngles(Scene(40, 40, 40), 40);
            AssertAngles(Scene(0, 0), 0);
            AssertAngles(Scene(0, 360), 0, 360);
            AssertAngles(Scene(360, 0), 360, 0);
            foreach (var sparseFullArc in new[] { Scene(0, 360), Scene(360, 0) })
                TestAssert.True(Frames(Evaluate(sparseFullArc)).All(f => (string)f.Status == "uncertain" && !(bool)f.Interpolated),
                    "Sparse full-turn endpoints cannot prove static geometry or an unambiguous interpolated arc.");
            var full = SyntheticCollisionFactory.Create("sweep-test");
            var frames = Frames(Evaluate(full));
            TestAssert.Equal(37, frames.Count, "Dense 0..360 fixture must keep ten-degree samples and both endpoints.");
            TestAssert.Equal(360.0, (double)frames.Last().Pose.GantryDegrees);
            var arc = Scene(3, 23); var original = arc.Poses[0];
            frames = Frames(Evaluate(arc));
            TestAssert.Equal(2, arc.Poses.Count); TestAssert.True(ReferenceEquals(original, arc.Poses[0]), "Input poses must stay untouched.");
            TestAssert.True((bool)frames[1].Interpolated, "Inserted milestone must be explicitly interpolated.");
            TestAssert.True(!(bool)frames[0].Interpolated, "Native endpoint must stay native.");
            TestAssert.Equal(frames.Count, frames.Select(f => (int)f.Pose.ControlPointIndex).Distinct().Count());
            var tilted = Scene(3, 23);
            foreach (var pose in tilted.Poses)
            {
                pose.PatientSupportAngleDegrees = 60;
                var s = pose.Source; pose.Source = new BeamPoint3D(s.X * 0.5, s.Y, s.X * Math.Sqrt(0.75));
            }
            var tiltedFrames = CollisionSweep.Evaluate(tilted, CancellationToken.None).Frames;
            TestAssert.Equal(4, tiltedFrames.Count);
            TestAssert.True(tiltedFrames.All(f => f.Pose.PatientSupportAngleDegrees == 60), "Couch must survive copied and interpolated rows.");
            TestAssert.True(tiltedFrames.Where(f => f.Interpolated).All(f => Math.Abs(f.Pose.Source.Z - Math.Sqrt(3) * f.Pose.Source.X) < 1e-6), "Interpolation must stay in the native tilted gantry plane.");
            tilted.Poses[1].PatientSupportAngleDegrees = 80;
            TestAssert.True(CollisionSweep.Evaluate(tilted, CancellationToken.None).Frames.All(f => !f.Interpolated && f.Status == "uncertain"), "Couch motion within an arc must not be interpolated as a gantry-only interval.");
        }

        public static void AmbiguousOrInconsistentArcsStayUncertain()
        {
            foreach (var scene in new[] { Scene(0, 170), Scene(0, 20), Scene(0, 20) })
            {
                if (scene.Poses[1].GantryDegrees == 20)
                    if (scene.Poses[0].Source.X == 0) scene.Poses[1].Source = new BeamPoint3D(0, -1000, 0);
                var frames = Frames(Evaluate(scene));
                TestAssert.Equal(2, frames.Count, "An ambiguous arc must not invent intermediate source geometry.");
                TestAssert.True(frames.All(f => (string)f.Status == "uncertain"), "Ambiguous intervals must remain yellow.");
            }
            var translated = Scene(0, 20); translated.Poses[1].Isocenter.X = 5;
            TestAssert.True(Frames(Evaluate(translated)).All(f => (string)f.Status == "uncertain"), "Changing isocenter must not be interpolated.");
            var stationarySource = Scene(9.99, 10.01); stationarySource.Poses[1].Source = stationarySource.Poses[0].Source;
            TestAssert.Equal(2, Frames(Evaluate(stationarySource)).Count, "A zero-angle source slerp must not generate NaN at a milestone.");
        }

        public static void CoverageAndCoordinateGatesPreserveProofs()
        {
            var scene = ProfileScene();
            TestAssert.Equal("pass", (string)Frames(Evaluate(scene))[0].Status);
            scene.Synthetic = false; scene.Profile.SyntheticOnly = false; scene.Profile.Commissioned = true;
            TestAssert.Equal("uncertain", (string)Frames(Evaluate(scene))[0].Status, "Unknown CT extent must not pass.");
            Set(scene, "CtCoverageKnown", true);
            TestAssert.Equal("pass", (string)Frames(Evaluate(scene))[0].Status);
            Set(scene, "ExternalTruncatedAtCtBoundary", true);
            TestAssert.Equal("uncertain", (string)Frames(Evaluate(scene))[0].Status);
            scene.Surfaces[0].Mesh = Box(-1, -7, -1, 1, -3, 1);
            TestAssert.Equal("hit", (string)Frames(Evaluate(scene))[0].Status, "A proven intersection outranks missing CT coverage.");
            scene.SupportedCoordinates = false; scene.NominalCoordinatesSupported = true;
            TestAssert.Equal("uncertain", (string)Frames(Evaluate(scene))[0].Status, "Unconfirmed coordinates cannot prove an authoritative hit.");
            TestAssert.True(Frames(Evaluate(scene))[0].BodyClearanceMm == null && Frames(Evaluate(scene))[0].TableClearanceMm == null,
                "A profile without confirmed coordinates is not drawable and must not supply unrelated distance numbers.");
            scene = ProfileScene(); scene.Surfaces.RemoveAt(1);
            var frame = Frames(Evaluate(scene))[0];
            TestAssert.Equal("uncertain", (string)frame.Status); TestAssert.Equal("uncertain", (string)frame.TableStatus);
            scene = ProfileScene(); scene.Surfaces[0].Mesh.TriangleIndices = scene.Surfaces[0].Mesh.TriangleIndices.Take(33).ToArray();
            TestAssert.Equal("uncertain", (string)Frames(Evaluate(scene))[0].Status, "An open external mesh cannot pass even in a synthetic fixture.");
            scene = ProfileScene(); scene.Surfaces.Add(new CollisionSurface { Label = "MISSING PART", Role = "External" });
            frame = Frames(Evaluate(scene))[0];
            TestAssert.True(frame.BodyLowerBoundMm == null && frame.BodyClearanceMm != null,
                "Incomplete role geometry invalidates its global lower bound while retaining measured samples from other surfaces.");
        }

        public static void SourceSignedDistancesAreNominalAndConservative()
        {
            var scene = SourceScene();
            // Head: axial front 100, radius 20, radial reach 200; source points toward +X.
            scene.Surfaces[0].Mesh = Box(89, -1, -1, 90, 1, 1);
            var frame = Frames(Evaluate(scene))[0];
            Near(10, (double?)frame.BodyClearanceMm, 0.01);
            scene.Profile = ProfileScene().Profile; scene.Profile.SyntheticOnly = false; scene.Profile.Commissioned = true;
            scene.SupportedCoordinates = false; scene.NominalCoordinatesSupported = true;
            frame = Frames(Evaluate(scene))[0]; Near(10, (double?)frame.BodyClearanceMm, 0.01);
            TestAssert.Equal("uncertain", (string)frame.Status, "A non-drawable commissioned profile must fall back to nominal source-model distances.");
            TestAssert.True(((string)frame.Reason).IndexOf("nominal source", StringComparison.OrdinalIgnoreCase) >= 0,
                "Distance provenance must match the source envelope drawn with nominal coordinates.");
            scene.Profile = null; scene.SupportedCoordinates = true;
            TestAssert.True((double)frame.BodyLowerBoundMm <= 10, "The lower bound cannot exceed true minimum clearance.");
            TestAssert.Equal("uncertain", (string)frame.Status, "Source models do not grant commissioning.");
            TestAssert.True(((string)frame.Reason).IndexOf("nominal", StringComparison.OrdinalIgnoreCase) >= 0, "Source estimates must be labelled nominal.");
            scene.Surfaces[0].Mesh = Box(110, 25, -1, 111, 26, 1);
            frame = Frames(Evaluate(scene))[0]; Near(5, (double?)frame.BodyClearanceMm, 0.01);
            scene.Surfaces[0].Mesh = Box(110, -1, -1, 111, 1, 1);
            frame = Frames(Evaluate(scene))[0]; TestAssert.True((double)frame.BodyClearanceMm < 0, "Inside points need negative signed distance.");
            TestAssert.Equal("hit", (string)frame.Status, "Confirmed source geometry can show a research-model hit.");
            scene.SupportedCoordinates = false; scene.NominalCoordinatesSupported = true;
            frame = Frames(Evaluate(scene))[0]; TestAssert.Equal("uncertain", (string)frame.Status);
            TestAssert.True((double)frame.BodyClearanceMm < 0, "Nominal estimates remain visible with explicit uncertainty.");
            scene.NominalCoordinatesSupported = false;
            TestAssert.True(Frames(Evaluate(scene))[0].BodyClearanceMm == null, "Unsupported coordinates must not produce a nominal number.");
            scene = SourceScene(); scene.SourceModel.ReachRadiusMm = double.NaN;
            TestAssert.True(Frames(Evaluate(scene))[0].BodyClearanceMm == null, "Invalid source model must fail closed.");
            scene = SourceScene(); scene.SourceModel.BoreLimitRadiusMm = 100;
            TestAssert.True(Frames(Evaluate(scene))[0].BodyClearanceMm == null, "Mixed head/bore models must fail closed.");

            scene = SourceScene(); scene.SourceModel.HeadObstacles.Add(new SourceHeadObstacle { FrontFaceFromIsoMm = 120, RadiusMm = 40 });
            scene.Surfaces[0].Mesh = PointTriangle(140, 0, 0);
            Near(-Math.Sqrt(800), (double?)Frames(Evaluate(scene))[0].BodyClearanceMm, 1e-8);
            scene = SourceScene(); scene.Surfaces[0].Mesh = PointTriangle(210, 0, 0);
            Near(10, (double?)Frames(Evaluate(scene))[0].BodyClearanceMm, 1e-8);
            scene.Surfaces[0].Mesh = PointTriangle(90, 30, 0);
            Near(Math.Sqrt(200), (double?)Frames(Evaluate(scene))[0].BodyClearanceMm, 1e-8);
            scene.Surfaces[0].Mesh = PointTriangle(200, 20, 0);
            Near(Math.Sqrt(40400) - 200, (double?)Frames(Evaluate(scene))[0].BodyClearanceMm, 1e-8);
        }

        public static void TriangleInteriorsAndHalcyonRoleDistances()
        {
            var scene = SourceScene();
            // All three vertices miss the head, while the triangle centroid lies inside it.
            scene.Surfaces[0].Mesh = new TargetMeshGeometry { Vertices = new List<BeamPoint3D> {
                new BeamPoint3D(110, -100, -100), new BeamPoint3D(110, 100, -100), new BeamPoint3D(110, 0, 200) }, TriangleIndices = new[] { 0, 1, 2 } };
            var frame = Frames(Evaluate(scene))[0];
            TestAssert.Equal("hit", (string)frame.Status, "A triangle-interior hit must be retained even on an open mesh.");
            TestAssert.True((double)frame.BodyLowerBoundMm <= (double)frame.BodyClearanceMm, "Sample value is an upper bound, not a minimum-surface claim.");
            scene = SourceScene(); scene.PatientPosition = "HFS";
            scene.SourceModel = new SourceCollisionModel { Kind = "HalcyonBoreSource", BoreWarningRadiusMm = 90, BoreLimitRadiusMm = 100,
                LeadingWarningFromIsoMm = 190, LeadingLimitFromIsoMm = 200 };
            scene.Surfaces[0].Mesh = Box(-3, -4, 189, 3, 4, 190);
            scene.Surfaces[1].Mesh = Box(-3, -4, 299, 3, 4, 300);
            frame = Frames(Evaluate(scene))[0]; Near(10, (double?)frame.BodyClearanceMm, 1e-6); Near(95, (double?)frame.TableClearanceMm, 1e-6);
            scene.Surfaces[0].Mesh = Box(99, -1, 0, 101, 1, 1);
            TestAssert.Equal("hit", (string)Frames(Evaluate(scene))[0].Status);
            scene.Surfaces[0].Mesh = PointTriangle(103, 0, 204);
            Near(-5, (double?)Frames(Evaluate(scene))[0].BodyClearanceMm, 1e-8);

            scene = SourceScene();
            // This triangle contains (110,0,0), but vertices and centroid all miss the thin head.
            scene.Surfaces[0].Mesh = new TargetMeshGeometry { Vertices = new List<BeamPoint3D> {
                new BeamPoint3D(110,-100,0), new BeamPoint3D(110,100,0), new BeamPoint3D(110,100,1000) }, TriangleIndices = new[] { 0,1,2 } };
            frame = Frames(Evaluate(scene))[0];
            TestAssert.True((double)frame.BodyClearanceMm > 0 && (double)frame.BodyLowerBoundMm < 0,
                "A missed interior overlap must leave a straddling conservative interval, not a clearance claim.");
            TestAssert.Equal("uncertain", (string)frame.Status);
            scene = SourceScene(); scene.Surfaces[0].Mesh = Box(10,-1,-1,11,1,1);
            scene.Surfaces[0].Mesh.Vertices.Add(new BeamPoint3D(110,0,0));
            frame = Frames(Evaluate(scene))[0];
            TestAssert.True((double)frame.BodyClearanceMm > 0 && (string)frame.Status == "uncertain",
                "Unreferenced mesh vertices are not triangle surface points and cannot prove a hit.");
        }

        public static void ProfileDistancesAndWorkloadCancellation()
        {
            var scene = ProfileScene(); var frame = Frames(Evaluate(scene))[0];
            Near(3, (double?)frame.BodyClearanceMm, 1e-8); Near(3, (double?)frame.BodyLowerBoundMm, 1e-8);
            scene.Profile.Kind = "RingBore"; scene.Profile.BoreRadiusMm = 10; scene.Profile.BoreHalfLengthMm = 5;
            scene.Surfaces[1].Mesh = Box(-2, -2, -1, 2, 2, 1);
            frame = Frames(Evaluate(scene))[0]; Near(10 - Math.Sqrt(2), (double?)frame.BodyClearanceMm, 1e-8);

            var rimScene = ProfileScene(); rimScene.Synthetic = false; Set(rimScene, "CtCoverageKnown", true);
            rimScene.Profile.SyntheticOnly = false; rimScene.Profile.Commissioned = true;
            rimScene.Profile.Kind = "RingBore"; rimScene.Profile.BoreRadiusMm = 10;
            rimScene.Profile.BoreHalfLengthMm = 5; rimScene.Profile.SafetyMarginMm = 2;
            rimScene.Surfaces[0].Mesh = Box(0,-0.1,4,0.2,0.1,6);
            foreach (var vertex in rimScene.Surfaces[0].Mesh.Vertices) if (vertex.Z == 6) vertex.X += 10.8;
            rimScene.Surfaces[1].Mesh = Box(-1,-1,-1,1,1,1);
            var rimFrame = Frames(Evaluate(rimScene))[0];
            Near(10 - Math.Sqrt(5.6 * 5.6 + 0.1 * 0.1), (double?)rimFrame.BodyClearanceMm, 1e-8);
            TestAssert.Equal("uncertain", (string)rimFrame.Status,
                "Clipped radial gap 4.399 mm cannot prove the 2-mm finite bore-end/rim margin for a sloping prism.");
            TestAssert.True(rimFrame.BodyLowerBoundMm == null && rimFrame.TableLowerBoundMm == null,
                "RingBore radial gaps are not conservative finite-end Euclidean clearance lower bounds.");
            TestAssert.True(((string)rimFrame.Reason).IndexOf("finite-end", StringComparison.OrdinalIgnoreCase) >= 0,
                "RingBore must explicitly disclose the missing finite-end/rim margin.");
            rimScene.Surfaces[0].Mesh = Box(-11,-0.1,0,11,0.1,1);
            TestAssert.Equal("hit", (string)Frames(Evaluate(rimScene))[0].Status,
                "A supported within-slab RingBore overlap proof remains a hit.");

            var cancellation = new CancellationTokenSource(); cancellation.Cancel();
            bool cancelled = false; try { Evaluate(scene, cancellation.Token); } catch (OperationCanceledException) { cancelled = true; }
            TestAssert.True(cancelled, "Cancellation must discard results, not return success.");
            scene = Scene(); scene.Poses = Enumerable.Range(0, 10001).Select(i => Pose(i, i % 360)).ToList();
            dynamic result = Evaluate(scene); TestAssert.Equal("uncertain", (string)result.Status);
            TestAssert.Equal(0, Frames(result).Count, "Over-budget inputs must stop before expensive geometry work.");
            scene = Scene(); scene.Poses = Enumerable.Range(0, 2001).Select(i => {
                var p = Pose(0, 0); p.BeamId = "BEAM-" + i; return p; }).ToList();
            TestAssert.Equal(0, Frames(Evaluate(scene)).Count, "Single-pose beams count toward the global frame budget.");
            scene = ProfileScene(); scene.Synthetic = false; scene.Profile.SyntheticOnly = false;
            Set(scene, "CtCoverageKnown", true);
            TestAssert.True(Frames(Evaluate(scene))[0].BodyClearanceMm == null, "A noncommissioned profile must never be enabled by sweep evaluation.");

            scene = SourceScene(); scene.NominalCoordinatesSupported = true;
            scene.Surfaces[0].Mesh = Box(89,-1,-1,90,1,1);
            scene.Profile = ProfileScene().Profile; scene.Profile.SyntheticOnly = false;
            frame = Frames(Evaluate(scene))[0]; Near(10, (double?)frame.BodyClearanceMm, 0.01);
            TestAssert.Equal("uncertain", (string)frame.Status, "An uncommissioned profile cannot suppress the valid nominal source envelope.");
            scene.Profile.Commissioned = true; scene.Profile.HeadRadiusMm = double.NaN;
            frame = Frames(Evaluate(scene))[0]; Near(10, (double?)frame.BodyClearanceMm, 0.01);
            TestAssert.Equal("uncertain", (string)frame.Status, "An invalid profile cannot promote or suppress source-model results.");
            scene.Profile.HeadRadiusMm = 1; Set(scene, "CtCoverageKnown", true);
            frame = Frames(Evaluate(scene))[0]; Near(83, (double?)frame.BodyClearanceMm, 1e-8);
            TestAssert.Equal("pass", (string)frame.Status, "A valid drawable profile retains precedence over nominal source geometry.");

            scene = SourceScene();
            scene.Surfaces[0].Mesh.TriangleIndices = Enumerable.Range(0, 500000).SelectMany(i => new[] { 0,1,2 }).ToArray();
            using (var duringWork = new CancellationTokenSource())
            {
                duringWork.CancelAfter(1); cancelled = false;
                try { Evaluate(scene, duringWork.Token); } catch (OperationCanceledException) { cancelled = true; }
                TestAssert.True(cancelled, "Cancellation must also be observed during bounded mesh preparation/evaluation.");
            }
        }

        public static void StationaryBoreReusesGeometryWithoutDroppingArcStates()
        {
            var scene = SourceScene();
            scene.SourceModel = new SourceCollisionModel { Kind = "HalcyonBoreSource", BoreWarningRadiusMm = 90,
                BoreLimitRadiusMm = 100, LeadingWarningFromIsoMm = 190, LeadingLimitFromIsoMm = 200 };
            scene.SupportedCoordinates = false; scene.NominalCoordinatesSupported = true;
            scene.CtCoverageKnown = true; scene.ExternalTruncatedAtCtBoundary = true;
            scene.Poses = Enumerable.Range(0, 37).Select(i => Pose(i, i * 10)).ToList();
            scene.Poses.AddRange(Enumerable.Range(0, 37).Select(i => { var p = Pose(i, 360 - i * 10); p.BeamId = "ARC2"; return p; }));
            // More than 10M triangle/angle pairs, but only one stationary bore geometry.
            scene.Surfaces[0].Mesh.TriangleIndices = Enumerable.Range(0, 150000).SelectMany(i => new[] { 0, 1, 2 }).ToArray();
            var result = CollisionSweep.Evaluate(scene, CancellationToken.None);
            TestAssert.Equal(74, result.Frames.Count, "A stationary bore must not lose all angle rows to repeated geometry work.");
            TestAssert.True(result.Frames.All(f => f.Status == "uncertain" && f.BodyClearanceMm.HasValue && f.TableClearanceMm.HasValue),
                "Every angle retains nominal distances and CT/coordinate uncertainty.");
            var first = result.Frames[0];
            TestAssert.True(result.Frames.All(f => f.BodyClearanceMm == first.BodyClearanceMm && f.BodyLowerBoundMm == first.BodyLowerBoundMm),
                "Identical bore geometry must retain identical conservative bounds.");
            // Another isocenter is a different envelope, even in a stationary ring.
            foreach (var pose in scene.Poses.Where(p => p.BeamId == "ARC2")) { pose.Isocenter.Z += 250; pose.Source.Z += 250; }
            scene.Surfaces[0].Mesh = Box(-1, -1, 240, 1, 1, 241);
            result = CollisionSweep.Evaluate(scene, CancellationToken.None);
            TestAssert.True(result.Frames[0].BodyClearanceMm != result.Frames[37].BodyClearanceMm, "Bore cache must include all isocenter coordinates.");
            foreach (var frame in result.Frames)
            {
                var saved = scene.Poses; scene.Poses = new List<CollisionPose> { frame.Pose };
                var independent = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single(); scene.Poses = saved;
                TestAssert.Equal(independent.BodyClearanceMm, frame.BodyClearanceMm);
                TestAssert.Equal(independent.TableLowerBoundMm, frame.TableLowerBoundMm);
            }
            // Dense repeated surfaces require bounded spatial pruning, not loss of every angle.
            scene.SourceModel = SourceScene().SourceModel;
            scene.Surfaces[0].Mesh.TriangleIndices = Enumerable.Range(0, 150000).SelectMany(i => new[] { 0, 1, 2 }).ToArray();
            var dense = CollisionSweep.Evaluate(scene, CancellationToken.None);
            TestAssert.Equal(74, dense.Frames.Count);
            var denseIndices = scene.Surfaces[0].Mesh.TriangleIndices;
            scene.Surfaces[0].Mesh.TriangleIndices = new[] { 0, 1, 2 };
            var simple = CollisionSweep.Evaluate(scene, CancellationToken.None);
            scene.Surfaces[0].Mesh.TriangleIndices = denseIndices;
            for (int i = 0; i < dense.Frames.Count; i++)
            {
                Near(simple.Frames[i].BodyClearanceMm.Value, dense.Frames[i].BodyClearanceMm, 1e-8);
                Near(simple.Frames[i].BodyLowerBoundMm.Value, dense.Frames[i].BodyLowerBoundMm, 1e-8);
                TestAssert.Equal("uncertain", dense.Frames[i].Status);
            }
        }

        public static void OpenBoreRadialReviewSeparatesCapturedClearanceFromClinicalAssurance()
        {
            TestAssert.NotNull(typeof(CollisionSweepFrame).GetProperty("RadialReview"), "Open-bore radial review needs its own explicitly limited result.");
            var scene = SourceScene(); scene.SupportedCoordinates = false; scene.NominalCoordinatesSupported = true;
            scene.CtCoverageKnown = true; scene.ExternalTruncatedAtCtBoundary = true;
            scene.SourceModel = new SourceCollisionModel { Kind = "HalcyonBoreSource", BoreWarningRadiusMm = 90,
                BoreLimitRadiusMm = 100, LeadingWarningFromIsoMm = 190, LeadingLimitFromIsoMm = 200 };
            scene.Surfaces[0].Mesh = Box(-3, -4, 299, 3, 4, 300); // Beyond the old axial threshold; open bore has no end wall.
            scene.Surfaces[1].Mesh = Box(-3, -4, 599, 3, 4, 600);
            dynamic frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            Near(95, (double?)frame.RadialReview.BodyClearanceMm, 1e-8);
            Near(95, (double?)frame.RadialReview.TableClearanceMm, 1e-8);
            TestAssert.Equal("clear", (string)frame.RadialReview.Status);
            var displayed = new ClearPlan.Presentation.ViewModels.CollisionSweepFrameViewModel((CollisionSweepFrame)frame);
            TestAssert.Equal("Full clear", displayed.StatusText);
            TestAssert.Equal("95.0", displayed.BodyDistanceText);
            TestAssert.Equal("uncertain", (string)frame.Status, "A positive captured radial distance does not manufacture clinical assurance or missing anatomy.");
            TestAssert.True(((string)frame.RadialReview.Scope).Contains("captured") && ((string)frame.RadialReview.Scope).Contains("outside CT"));
            scene.Surfaces[0].Mesh = Box(99, -1, 299, 101, 1, 300);
            frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            TestAssert.Equal("model-hit", (string)frame.RadialReview.Status);
            scene.Surfaces[0].Mesh = PointTriangle(0, 0, 0);
            frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            TestAssert.Equal("uncertain", (string)frame.RadialReview.Status, "Degenerate geometry must not become radial clear.");
            TestAssert.True(((string)frame.RadialReview.Scope).Contains("degenerate"), "Invalid geometry reason must remain visible.");
            scene.Surfaces[0].Mesh = Box(-3, -4, 299, 3, 4, 300);
            scene.Surfaces[0].Mesh.Vertices.Add(new BeamPoint3D(-3, -4, 299)); // Duplicate WPF index, same real face vertex.
            scene.Surfaces[0].Mesh.TriangleIndices = scene.Surfaces[0].Mesh.TriangleIndices.Concat(new[] { 8, 8, 8 }).ToArray();
            frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            TestAssert.Equal("clear", (string)frame.RadialReview.Status, "Redundant zero-area triangles at real surface vertices must not discard the complete captured surface.");
            Near(95, (double?)frame.RadialReview.BodyClearanceMm, 1e-8);
            scene.Surfaces[0].Mesh.Vertices[8].X = 0; // Dangling degenerate geometry not represented by any proper face.
            frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            TestAssert.Equal("uncertain", (string)frame.RadialReview.Status, "Detached unsupported primitives remain uncertain.");
            scene.Surfaces[0].Mesh = null;
            frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            TestAssert.Equal("uncertain", (string)frame.RadialReview.Status);
            scene.NominalCoordinatesSupported = false;
            frame = CollisionSweep.Evaluate(scene, CancellationToken.None).Frames.Single();
            TestAssert.True(frame.RadialReview == null, "Unknown native positioning cannot yield a radial clearance result.");
        }

        public static void SpatialPruningKeepsBoundsAndHonorsLimits()
        {
            var type = typeof(CollisionScene).Assembly.GetType("ClearPlan.Core.Collision.CollisionSweepDistance");
            var solidType = type.GetNestedType("SourceSolid", BindingFlags.NonPublic);
            var budgetType = type.GetNestedType("DistanceBudget", BindingFlags.NonPublic);
            const BindingFlags statics = BindingFlags.NonPublic | BindingFlags.Static;
            const BindingFlags instance = BindingFlags.NonPublic | BindingFlags.Instance;
            foreach (double plane in new[] { 90.0, 110.0 }) foreach (bool rotated in new[] { false, true })
            {
                var scene = SourceScene(); var mesh = new TargetMeshGeometry { Vertices = new List<BeamPoint3D>() };
                Func<BeamPoint3D,BeamPoint3D> transform = p => rotated ? new BeamPoint3D(-p.Z+20,p.Y-30,p.X+50) : p;
                var indices = new List<int>();
                for (int y=0;y<=40;y++) for (int z=0;z<=40;z++) mesh.Vertices.Add(transform(new BeamPoint3D(plane,(y-20)*10,(z-20)*10)));
                for (int y=0;y<40;y++) for (int z=0;z<40;z++) { int a=y*41+z; indices.AddRange(new[] { a,a+41,a+1,a+1,a+41,a+42 }); }
                mesh.TriangleIndices=indices.ToArray(); scene.Surfaces[0].Mesh=mesh;
                scene.Poses[0].Source=transform(scene.Poses[0].Source); scene.Poses[0].Isocenter=transform(scene.Poses[0].Isocenter);
                var prepared=type.GetMethod("Prepare",statics).Invoke(null,new object[] {scene.Surfaces[0],CancellationToken.None});
                var solid=solidType.GetMethod("Create",statics).Invoke(null,new object[] {scene.SourceModel});
                var distance=solidType.GetMethod("AtPose",instance).Invoke(solid,new object[] {scene.Poses[0]});
                var budget=Activator.CreateInstance(budgetType,true);
                var args=new object[] {prepared,distance,CancellationToken.None,budget,null,null};
                type.GetMethod("SampleSpatial",statics).Invoke(null,args);
                var exhaustive=new object[] {prepared,distance,CancellationToken.None,null,null};
                type.GetMethod("Sample",statics).Invoke(null,exhaustive);
                double expected=plane==90 ? 10 : -10;
                Near(expected,(double?)args[5],1e-8); Near(expected,(double?)exhaustive[4],1e-8);
                TestAssert.True((double?)args[4] <= expected+1e-8, "Spatial lower bound must remain conservative under rigid transforms.");
                long used=CollisionSweep.MaximumTrianglePosePairs-(long)budgetType.GetField("remaining",instance).GetValue(budget);
                TestAssert.True(used>0 && used<mesh.Vertices.Count+3200, "Unique triangles must exercise pruning rather than only duplicate reuse.");
                budgetType.GetField("remaining",instance).SetValue(budget,10L);
                bool stopped=false;
                try { type.GetMethod("SampleSpatial",statics).Invoke(null,args); }
                catch (TargetInvocationException error) { stopped=error.InnerException is ArgumentException; }
                TestAssert.True(stopped, "Traversal must stop when its actual distance budget is exhausted.");
                using (var cancellation=new CancellationTokenSource())
                {
                    cancellation.Cancel(); args[2]=cancellation.Token; args[3]=Activator.CreateInstance(budgetType,true); stopped=false;
                    try { type.GetMethod("SampleSpatial",statics).Invoke(null,args); }
                    catch (TargetInvocationException error) { stopped=error.InnerException is OperationCanceledException; }
                    TestAssert.True(stopped, "A cached hierarchy must still honor cancellation.");
                }
            }
        }

        private static object Evaluate(CollisionScene scene, CancellationToken token = default(CancellationToken))
        {
            DetachedContractExists();
            try { return typeof(CollisionScene).Assembly.GetType("ClearPlan.Core.Collision.CollisionSweep").GetMethod("Evaluate").Invoke(null, new object[] { scene, token }); }
            catch (TargetInvocationException error) { throw error.InnerException; }
        }
        private static List<dynamic> Frames(object result) { return ((IEnumerable)((dynamic)result).Frames).Cast<dynamic>().ToList(); }
        private static void Set(CollisionScene scene, string property, object value) { typeof(CollisionScene).GetProperty(property).SetValue(scene, value); }
        private static void AssertAngles(CollisionScene scene, params double[] expected)
        {
            var actual = Frames(Evaluate(scene)).Select(f => (double)f.Pose.GantryDegrees).ToArray();
            TestAssert.Equal(string.Join(",", expected), string.Join(",", actual));
        }
        private static void Near(double expected, double? actual, double tolerance)
        { TestAssert.True(actual.HasValue && Math.Abs(expected - actual.Value) <= tolerance, "Expected " + expected + " mm, got " + actual + "."); }
        private static CollisionScene Scene(params double[] angles)
        {
            var scene = ProfileScene(); scene.Poses = (angles.Length == 0 ? new[] { 0.0 } : angles).Select((a, i) => Pose(i, a)).ToList(); return scene;
        }
        private static CollisionPose Pose(int index, double angle)
        { var a = angle * Math.PI / 180; return new CollisionPose { BeamId = "TEST", ControlPointIndex = index, GantryDegrees = angle,
            Isocenter = new BeamPoint3D(0, 0, 0), Source = new BeamPoint3D(1000 * Math.Sin(a), -1000 * Math.Cos(a), 0) }; }
        private static CollisionScene ProfileScene()
        {
            return new CollisionScene { Synthetic = true, SupportedCoordinates = true, PatientPosition = "HFS",
                Profile = new CollisionProfile { Kind = "CArmSphere", MachineId = "TEST", Revision = "test", CommissionedBy = "synthetic", Evidence = "test-only geometry",
                    SyntheticOnly = true, HeadCenterFromIsoMm = 5, HeadRadiusMm = 1, SafetyMarginMm = 0 },
                Poses = new List<CollisionPose> { Pose(0, 0) }, Surfaces = new List<CollisionSurface> {
                    new CollisionSurface { Label = "BODY", Role = "External", Mesh = Box(-1,-1,-1,1,1,1) },
                    new CollisionSurface { Label = "TABLE", Role = "Support", Mesh = Box(-1,10,-1,1,12,1) } } };
        }
        private static CollisionScene SourceScene()
        {
            var scene = ProfileScene(); scene.Synthetic = false; scene.Profile = null;
            scene.Poses[0].Source = new BeamPoint3D(1000, 0, 0);
            scene.SourceModel = new SourceCollisionModel { Kind = "TrueBeamHeadSource", ReachRadiusMm = 200,
                HeadObstacles = new List<SourceHeadObstacle> { new SourceHeadObstacle { FrontFaceFromIsoMm = 100, RadiusMm = 20 } } };
            return scene;
        }
        private static TargetMeshGeometry Box(double x0, double y0, double z0, double x1, double y1, double z1)
        {
            return new TargetMeshGeometry { Vertices = new List<BeamPoint3D> { new BeamPoint3D(x0,y0,z0), new BeamPoint3D(x1,y0,z0),
                new BeamPoint3D(x1,y1,z0), new BeamPoint3D(x0,y1,z0), new BeamPoint3D(x0,y0,z1), new BeamPoint3D(x1,y0,z1),
                new BeamPoint3D(x1,y1,z1), new BeamPoint3D(x0,y1,z1) },
                TriangleIndices = new[] { 0,2,1,0,3,2,4,5,6,4,6,7,0,1,5,0,5,4,1,2,6,1,6,5,2,3,7,2,7,6,3,0,4,3,4,7 } };
        }
        private static TargetMeshGeometry PointTriangle(double x, double y, double z)
        {
            return new TargetMeshGeometry { Vertices = new List<BeamPoint3D> { new BeamPoint3D(x,y,z), new BeamPoint3D(x,y,z),
                new BeamPoint3D(x,y,z) }, TriangleIndices = new[] { 0,1,2 } };
        }
    }
}
