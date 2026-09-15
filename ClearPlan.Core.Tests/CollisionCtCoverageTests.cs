using System;
using ClearPlan.Core.Collision;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Tests
{
    internal static class CollisionCtCoverageTests
    {
        public static void ImageBoundsDetectCappedExternalAndInvalidMetadata()
        {
            var type = typeof(CollisionScene).Assembly.GetType("ClearPlan.Core.Collision.CollisionCtCoverage");
            TestAssert.NotNull(type, "Native image coverage must be checked independently of mesh closure.");
            var method = type.GetMethod("Apply");
            TestAssert.NotNull(method);
            var scene = SyntheticCollisionFactory.Create("coverage");
            var axes = new[] { new BeamPoint3D(1,0,0), new BeamPoint3D(0,1,0), new BeamPoint3D(0,0,1) };
            method.Invoke(null, new object[] { scene, new BeamPoint3D(-500,-500,-500), axes, new[] { 1.0,1.0,1.0 }, new[] {1001,1001,1001} });
            TestAssert.Equal(true, scene.GetType().GetProperty("CtCoverageKnown").GetValue(scene));
            TestAssert.Equal(false, scene.GetType().GetProperty("ExternalTruncatedAtCtBoundary").GetValue(scene));
            method.Invoke(null, new object[] { scene, new BeamPoint3D(-180,-500,-500), axes, new[] { 1.0,1.0,1.0 }, new[] {361,1001,1001} });
            TestAssert.Equal(true, scene.GetType().GetProperty("ExternalTruncatedAtCtBoundary").GetValue(scene));
            // A reversed slice axis uses image coordinates, not patient-z sign assumptions.
            axes[2] = new BeamPoint3D(0,0,-1);
            method.Invoke(null, new object[] { scene, new BeamPoint3D(-500,-500,280), axes, new[] { 1.0,1.0,2.0 }, new[] {1001,1001,281} });
            TestAssert.Equal(true, scene.GetType().GetProperty("ExternalTruncatedAtCtBoundary").GetValue(scene));
            method.Invoke(null, new object[] { scene, new BeamPoint3D(-500,-500,-500), axes, new[] { 1.0,1.0,double.NaN }, new[] {1001,1001,1001} });
            TestAssert.Equal(false, scene.GetType().GetProperty("CtCoverageKnown").GetValue(scene));
        }
    }
}
