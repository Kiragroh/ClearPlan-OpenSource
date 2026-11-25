using System;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Simulation
{
    public static class SyntheticDrrFactory
    {
        private static readonly Lazy<CtVolume> Phantom=new Lazy<CtVolume>(CreatePhantom);

        public static BeamEyeViewImage Create(ReviewBeamAnalysis beam,ReviewControlPointSample cp)
        {
            if(beam==null) throw new ArgumentNullException("beam");
            if(cp==null) throw new ArgumentNullException("cp");
            var frame=CreateHfsFrame(cp.GantryAngleDegrees,cp.CollimatorAngleDegrees,cp.PatientSupportAngleDegrees,1000);
            var image=CtDrrProjector.Project(Phantom.Value,frame,224,220,CancellationToken.None);
            image.Synthetic=true; image.SourceStatus="synthetic"; image.ControlPointIndex=cp.Index;
            image.GantryAngleDegrees=cp.GantryAngleDegrees; image.CollimatorAngleDegrees=cp.CollimatorAngleDegrees;
            image.BldToDisplayRotationDegrees=cp.CollimatorAngleDegrees % 360;
            image.PatientSupportAngleDegrees=cp.PatientSupportAngleDegrees;
            image.ProjectionDescription="Original deterministic head/torso ellipsoid phantom with bone, lungs and air; no patient data. " +
                "Illustrative HFS, couch 0, SAD 1000 mm frame; positive collimator angle uses the declared synthetic beam-coordinate rotation. " +
                CtDrrProjector.MethodDescription;
            return image;
        }

        /// <summary>
        /// Explicit synthetic/offline HFS convention only: G0 source anterior (-Y), G90 source +X,
        /// positive beam X patient-left and beam Y superior at G0/C0. Positive C is a positive 2D
        /// coordinate rotation (world X/Z -> BEV -Z/X at C90). Couch rotations are not supported.
        /// Live ESAPI must supply its vendor-derived frame instead of using this helper.
        /// </summary>
        public static BeamProjectionFrame CreateHfsFrame(double gantryDegrees,double collimatorDegrees,double patientSupportDegrees,double sadMm)
        {
            if(!Finite(gantryDegrees) || !Finite(collimatorDegrees) || !Finite(patientSupportDegrees) ||
                !Finite(sadMm) || sadMm<100 || sadMm>5000 || Math.Abs(Math.IEEERemainder(patientSupportDegrees,360))>1e-6)
                throw new ArgumentException("Synthetic HFS frame requires finite angles, couch 0 and a supported source distance.");
            double gantry=(gantryDegrees%360)*Math.PI/180;
            return MeshTargetProjector.CreateFrame(new BeamPoint3D(),
                new BeamPoint3D(sadMm*Math.Sin(gantry),-sadMm*Math.Cos(gantry),0),
                new BeamPoint3D(0,-sadMm,0),new BeamPoint3D(sadMm,0,0),(collimatorDegrees%360)*Math.PI/180);
        }

        private static CtVolume CreatePhantom()
        {
            var volume=new CtVolume { SizeX=128,SizeY=96,SizeZ=144,SpacingX=3,SpacingY=3,SpacingZ=3,
                Origin=new BeamPoint3D(-190.5,-142.5,-214.5),XAxis=new BeamPoint3D(1,0,0),
                YAxis=new BeamPoint3D(0,1,0),ZAxis=new BeamPoint3D(0,0,1),HounsfieldUnits=new float[128*96*144] };
            for(int z=0;z<volume.SizeZ;z++) for(int y=0;y<volume.SizeY;y++) for(int x=0;x<volume.SizeX;x++)
                volume.HounsfieldUnits[(z*volume.SizeY+y)*volume.SizeX+x]=PhantomHu(
                    volume.Origin.X+x*3,volume.Origin.Y+y*3,volume.Origin.Z+z*3);
            return volume;
        }

        private static float PhantomHu(double x,double y,double z)
        {
            double torso=Ellipsoid(x,y,z+65,135,95,160);
            double skull=Ellipsoid(x,y,z-147,66,78,63);
            bool neck=z>62 && z<132 && Ellipsoid(x,y,0,34,40,1)<1;
            if(torso>1 && skull>1 && !neck) return -1000;
            float hu=35;
            if(skull<=1)
            {
                hu=Ellipsoid(x,y,z-147,60,72,57)>1 ? 1050 : 40;
                if(Ellipsoid(x-21,y+54,z-139,15,20,13)<1 || Ellipsoid(x+21,y+54,z-139,15,20,13)<1) hu=-850;
            }
            if(torso<=1)
            {
                if(Ellipsoid(x-49,y+2,z+35,43,62,95)<1 || Ellipsoid(x+49,y+2,z+35,43,62,95)<1) hu=-760;
                // A small off-axis mediastinal ellipsoid makes left/right orientation visibly asymmetric.
                if(Ellipsoid(x-20,y+28,z+80,37,35,46)<1) hu=70;
                double ribRadius=x*x/(135*135)+y*y/(95*95);
                if(z>-150 && z<65 && Math.Abs(Math.IEEERemainder(z+10,22))<2.4 && ribRadius>0.69 && ribRadius<0.87) hu=1100;
            }
            if(z>-194 && z<102 && Ellipsoid(x,y-52,0,12,13,1)<1)
                hu=Ellipsoid(x,y-52,0,4,5,1)<1 ? 25 : 1200;
            return hu;
        }

        private static double Ellipsoid(double x,double y,double z,double rx,double ry,double rz)
        { return x*x/(rx*rx)+y*y/(ry*ry)+z*z/(rz*rz); }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
    }
}
