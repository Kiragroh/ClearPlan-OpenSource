using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Tests
{
    internal static class MeshTargetProjectionTests
    {
        public static void AnalyticFramesAndPerspective()
        {
            var iso = new BeamPoint3D(0,0,0);
            var zero = new BeamPoint3D(0,-1000,0);
            var ninety = new BeamPoint3D(1000,0,0);
            var front = MeshTargetProjector.CreateFrame(iso,zero,zero,ninety,0);
            TestAssert.NotNull(front, "Vendor-source projection frame is not implemented.");
            var projected = MeshTargetProjector.ProjectPoint(new BeamPoint3D(10,0,20),front);
            Near(10,projected.X); Near(20,projected.Y);
            projected = MeshTargetProjector.ProjectPoint(new BeamPoint3D(10,-500,20),front);
            Near(20,projected.X); Near(40,projected.Y);
            var side = MeshTargetProjector.CreateFrame(iso,ninety,zero,ninety,0);
            projected = MeshTargetProjector.ProjectPoint(new BeamPoint3D(0,10,20),side);
            Near(10,projected.X); Near(20,projected.Y);
            var rotated = MeshTargetProjector.CreateFrame(iso,zero,zero,ninety,Math.PI/2);
            projected = MeshTargetProjector.ProjectPoint(new BeamPoint3D(10,0,20),rotated);
            Near(-20,projected.X); Near(10,projected.Y);
            // A tilted vendor gantry-plane basis, rather than a hard-coded patient Z axis.
            var tilted = MeshTargetProjector.CreateFrame(iso,new BeamPoint3D(0,-1000,0),
                new BeamPoint3D(0,-1000,0),new BeamPoint3D(0,0,-1000),0);
            projected = MeshTargetProjector.ProjectPoint(new BeamPoint3D(20,0,-10),tilted);
            Near(10,projected.X); Near(20,projected.Y);
            TestAssert.Throws<ArgumentException>(() => MeshTargetProjector.ProjectPoint(new BeamPoint3D(0,-1000,0),front));
        }

        public static void TriangleUnionNotXorAndNoConvexHull()
        {
            var frame = MeshTargetProjector.CreateFrame(new BeamPoint3D(0,0,0),new BeamPoint3D(0,-1000,0),
                new BeamPoint3D(0,-1000,0),new BeamPoint3D(1000,0,0),0);
            var mesh = new TargetMeshGeometry { Vertices = new List<BeamPoint3D> {
                new BeamPoint3D(-10,0,-10),new BeamPoint3D(10,0,-10),new BeamPoint3D(10,0,10),new BeamPoint3D(-10,0,10) },
                TriangleIndices = new[] { 0,1,2,0,2,3, 0,1,2,0,2,3 } };
            var strips = MeshTargetProjector.Project(mesh,frame,0.625);
            Near(400,Area(strips)); // Duplicate front/back faces must not cancel.
            var outlines = new List<List<BeamPoint>> { Rectangle(-10,-10,10,10) };
            Near(0,MeshTargetProjector.SymmetricDifferenceFraction(strips,MeshTargetProjector.RasterizeOutlines(outlines,0.625)));
            mesh.Vertices.AddRange(new[] { new BeamPoint3D(20,0,-10),new BeamPoint3D(30,0,-10),new BeamPoint3D(30,0,10),new BeamPoint3D(20,0,10) });
            mesh.TriangleIndices = new[] { 0,1,2,0,2,3, 4,5,6,4,6,7 };
            strips = MeshTargetProjector.Project(mesh,frame,0.625);
            Near(600,Area(strips)); // Gap between disconnected targets is not filled.
            TestAssert.True(MeshTargetProjector.SymmetricDifferenceFraction(strips,
                MeshTargetProjector.RasterizeOutlines(new List<List<BeamPoint>> { Rectangle(-10,-10,30,10) },0.625))>0.2);
        }

        public static void PlanPamUsesUnionRows()
        {
            var strips=MeshTargetProjector.RasterizeOutlines(new List<List<BeamPoint>> { Rectangle(-10,-10,10,10) },0.625);
            var plan=new ReviewPlanAnalysis { TargetStructureId="SYNTHETIC_TARGET",DosePerFractionGy=2 };
            var beam=new ReviewBeamAnalysis { MetersetMu=100 };
            for(int i=0;i<2;i++) beam.ControlPoints.Add(new ReviewControlPointSample {
                Index=i,CumulativeMetersetWeight=i,TargetProjectionStrips=strips,
                Aperture=new ApertureGeometry { Jaws=new ApertureRectangle(-10,-10,0,10) } });
            plan.Beams.Add(beam);
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.True(plan.Pam.HasValue,"Projected target union rows must feed PAM.");
            Near(0.5,plan.Pam.Value);Near(4,beam.ControlPoints[0].TargetAreaCm2.Value);
            beam.ControlPoints[0].Aperture=null;
            PlanAnalysisCalculator.Calculate(plan);
            TestAssert.False(plan.Pam.HasValue);
            Near(4,beam.ControlPoints[0].TargetAreaCm2.Value);
        }

        public static void NativeRotationUsesCoordinatesNotAngleGuess()
        {
            var original=new List<List<BeamPoint>> { Rectangle(-10,-5,20,15) };
            var rotated=original.Select(l=>l.Select(p=>new BeamPoint(-p.Y,p.X)).ToList()).ToList();
            Near(Math.PI/2,MeshTargetProjector.NativeCoordinateRotation(original,rotated));
            rotated[0][0].X+=2;
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.NativeCoordinateRotation(original,rotated));
        }

        public static void NativeRotationAcceptsIndependentContourSampling()
        {
            var original = new List<List<BeamPoint>> { Rectangle(-10,-5,20,15) };
            double a=-40*Math.PI/180, c=Math.Cos(a),s=Math.Sin(a);
            var dense=original[0].SelectMany((p,i)=>new[] { p,
                new BeamPoint((p.X+original[0][(i+1)%4].X)/2,(p.Y+original[0][(i+1)%4].Y)/2) });
            var rotated = new List<List<BeamPoint>> { dense.Reverse().Select(p=>new BeamPoint(c*p.X-s*p.Y,s*p.X+c*p.Y)).ToList() };
            Near(a,MeshTargetProjector.NativeCoordinateRotation(original,rotated,40));
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.NativeCoordinateRotation(original,rotated,20));
            rotated[0][0].X+=2;
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.NativeCoordinateRotation(original,rotated,40));
            var symmetric=new List<List<BeamPoint>> { Rectangle(-10,-10,10,10) };
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.NativeCoordinateRotation(symmetric,symmetric,90));
            Near(0,MeshTargetProjector.NativeCoordinateRotation(symmetric,symmetric,0));
        }

        public static void CancellationAndInvalidFramesAreExplicit()
        {
            var frame=MeshTargetProjector.CreateFrame(new BeamPoint3D(0,0,0),new BeamPoint3D(0,-1000,0),
                new BeamPoint3D(0,-1000,0),new BeamPoint3D(1000,0,0),0);
            var mesh=new TargetMeshGeometry { Vertices=new List<BeamPoint3D> {
                new BeamPoint3D(-10,0,-10),new BeamPoint3D(10,0,-10),new BeamPoint3D(0,0,10) },TriangleIndices=new[] {0,1,2} };
            using(var canceled=new System.Threading.CancellationTokenSource())
            {
                canceled.Cancel();
                TestAssert.Throws<OperationCanceledException>(()=>MeshTargetProjector.Project(mesh,frame,0.625,canceled.Token));
            }
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.Project(mesh,frame,0));
            frame.XAxis=new BeamPoint3D(2,0,0);
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.Project(mesh,frame,0.625));
            TestAssert.Throws<ArgumentException>(()=>MeshTargetProjector.CreateFrame(new BeamPoint3D(0,0,0),new BeamPoint3D(0,-1000,0),
                new BeamPoint3D(0,-1000,0),new BeamPoint3D(0,-1000,0),0));
        }

        private static List<BeamPoint> Rectangle(double x1,double y1,double x2,double y2)
        { return new List<BeamPoint> { new BeamPoint(x1,y1),new BeamPoint(x2,y1),new BeamPoint(x2,y2),new BeamPoint(x1,y2) }; }
        private static double Area(List<ApertureRectangle> rows) { return rows.Sum(r=>(r.X2-r.X1)*(r.Y2-r.Y1)); }
        private static void Near(double a,double b) { TestAssert.True(Math.Abs(a-b)<1e-8,"Expected " + a + "; got " + b); }
    }
}
