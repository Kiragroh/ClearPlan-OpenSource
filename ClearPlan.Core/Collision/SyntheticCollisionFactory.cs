using System;
using System.Collections.Generic;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    public static class SyntheticCollisionFactory
    {
        public static CollisionScene Create(string planKey, string kind = "CArmSphere")
        {
            if (kind != "CArmSphere" && kind != "RingBore") throw new ArgumentException("Unknown synthetic envelope kind.");
            var scene = new CollisionScene { PlanKey = planKey, Synthetic = true, SupportedCoordinates = true,
                GeometryReason = "SYNTHETIC ellipsoid and table only; arbitrary demonstration dimensions, not a device specification.",
                Profile = new CollisionProfile { MachineId = "SYNTHETIC-" + kind, Kind = kind, Revision = "synthetic-1",
                    CommissionedBy = "Synthetic fixture only", Evidence = "Arbitrary demonstration geometry; no clinical device dimensions or commissioning.",
                    Commissioned = false, SyntheticOnly = true, HeadCenterFromIsoMm = 330, HeadRadiusMm = 120,
                    BoreRadiusMm = 240, BoreHalfLengthMm = 220, SafetyMarginMm = 25 } };
            scene.Surfaces.Add(new CollisionSurface { Label = "SYNTHETIC External", Role = "External", Mesh = Ellipsoid(180,105,280) });
            scene.Surfaces.Add(new CollisionSurface { Label = "SYNTHETIC Support", Role = "Support", Mesh = Table() });
            for (int angle = 0; angle <= 360; angle += 5)
            {
                double radians = angle*Math.PI/180;
                scene.Poses.Add(new CollisionPose { BeamId = "SYNTHETIC ARC", ControlPointIndex = angle/5, GantryDegrees = angle,
                    Isocenter = new BeamPoint3D(0,0,0), Source = new BeamPoint3D(1000*Math.Sin(radians),-1000*Math.Cos(radians),0) });
            }
            return scene;
        }

        private static TargetMeshGeometry Ellipsoid(double x, double y, double z)
        {
            const int latitude = 16, longitude = 32;
            var vertices = new List<BeamPoint3D> { new BeamPoint3D(0,0,-z) };
            for (int ring = 1; ring < latitude; ring++)
            {
                double a = -Math.PI/2 + ring*Math.PI/latitude;
                for (int j = 0; j < longitude; j++)
                {
                    double b = j*2*Math.PI/longitude;
                    vertices.Add(new BeamPoint3D(x*Math.Cos(a)*Math.Cos(b),y*Math.Cos(a)*Math.Sin(b),z*Math.Sin(a)));
                }
            }
            int north = vertices.Count; vertices.Add(new BeamPoint3D(0,0,z));
            var triangles = new List<int>();
            for (int j = 0; j < longitude; j++)
            {
                int next = (j+1)%longitude;
                triangles.AddRange(new[] { 0,1+next,1+j });
                for (int ring = 0; ring < latitude-2; ring++)
                {
                    int a = 1+ring*longitude+j, b = 1+ring*longitude+next;
                    triangles.AddRange(new[] { a,b,b+longitude,a,b+longitude,a+longitude });
                }
                int last = 1+(latitude-2)*longitude;
                triangles.AddRange(new[] { last+j,last+next,north });
            }
            return new TargetMeshGeometry { Vertices = vertices, TriangleIndices = triangles.ToArray() };
        }

        private static TargetMeshGeometry Table()
        {
            return new TargetMeshGeometry { Vertices = new List<BeamPoint3D> {
                new BeamPoint3D(-210,125,-320),new BeamPoint3D(210,125,-320),new BeamPoint3D(210,150,-320),new BeamPoint3D(-210,150,-320),
                new BeamPoint3D(-210,125,320),new BeamPoint3D(210,125,320),new BeamPoint3D(210,150,320),new BeamPoint3D(-210,150,320) },
                TriangleIndices = new[] { 0,2,1,0,3,2,4,5,6,4,6,7,0,1,5,0,5,4,1,2,6,1,6,5,2,3,7,2,7,6,3,0,4,3,4,7 } };
        }
    }
}
