using System;
using System.Linq;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Simulation;
using Newtonsoft.Json;

namespace ClearPlan.Core.Tests
{
    internal static class DrrProjectionTests
    {
        public static void RunAll()
        {
            WaterSlabUsesFiniteVoxelSupport();
            AirHasZeroAttenuation();
            AsymmetryPreservesBeamAxesAndTopRow();
            SourceDistanceProducesDivergentMagnification();
            RotatedVolumeAndTranslatedFramePreservePath();
            InvalidGeometryAndBudgetsAreRejected();
            CancellationIsObserved();
            PixelsAndCtAreNeverSerialized();
            HfsCardinalFramesAndCollimatorAreExplicit();
            SyntheticSingleAndDualLayerHaveAnatomicalContrast();
        }

        public static void WaterSlabUsesFiniteVoxelSupport()
        {
            var volume = Uniform(21,10,21,2,new BeamPoint3D(-20,-9,-20),0);
            Near(20,CtDrrProjector.IntegrateRay(volume,Front(),0,0,CancellationToken.None),1e-8);
            Near(20*Math.Sqrt(1+0.01*0.01),
                CtDrrProjector.IntegrateRay(volume,Front(),10,0,CancellationToken.None),1e-8);
            TestAssert.Equal(0.0,CtDrrProjector.IntegrateRay(volume,Front(),80,0,CancellationToken.None));
            // A one-voxel-wide volume still has finite support rather than zero thickness.
            var one = Uniform(1,1,1,4,new BeamPoint3D(),0);
            Near(4,CtDrrProjector.IntegrateRay(one,Front(),0,0,CancellationToken.None),1e-8);
        }

        public static void AirHasZeroAttenuation()
        {
            var air = Uniform(5,5,5,2,new BeamPoint3D(-4,-4,-4),-1000);
            Near(0,CtDrrProjector.IntegrateRay(air,Front(),0,0,CancellationToken.None),0);
            var image = CtDrrProjector.Project(air,Front(),17,20,CancellationToken.None);
            TestAssert.NotNull(image,"DRR image is not implemented.");
            TestAssert.True(image.GrayscalePixels.All(p=>p==0),"Air must remain black after windowing.");
            TestAssert.Equal(17,image.WidthPixels);
            TestAssert.Equal(17,image.HeightPixels);
            TestAssert.Equal(20.0,image.ExtentMm);
            TestAssert.True(image.ProjectionDescription.Contains("not dose"));
        }

        public static void AsymmetryPreservesBeamAxesAndTopRow()
        {
            var volume = Uniform(65,65,65,2,new BeamPoint3D(-64,-64,-64),-1000);
            for(int z=42;z<=52;z++) for(int y=27;y<=37;y++) for(int x=38;x<=46;x++)
                volume.HounsfieldUnits[(z*65+y)*65+x]=2000;
            var image = CtDrrProjector.Project(volume,Front(),65,64,CancellationToken.None);
            TestAssert.NotNull(image,"DRR image is not implemented.");
            double weightedX=0,weightedY=0,total=0;
            for(int y=0;y<65;y++) for(int x=0;x<65;x++)
            {
                double value=image.GrayscalePixels[y*65+x];
                total+=value; weightedX+=value*x; weightedY+=value*y;
            }
            TestAssert.True(total>0,"Asymmetric object must be visible.");
            TestAssert.True(weightedX/total>38,"Positive BEV-X must be on the right.");
            TestAssert.True(weightedY/total<24,"Positive BEV-Y must be at the top.");
            TestAssert.Equal((byte)0,image.GrayscalePixels[47*65+22]);
        }

        public static void SourceDistanceProducesDivergentMagnification()
        {
            var near = Uniform(21,2,21,1,new BeamPoint3D(0,-500.5,-10),-1000);
            for(int z=0;z<21;z++) for(int y=0;y<2;y++) for(int x=9;x<=11;x++)
                near.HounsfieldUnits[(z*2+y)*21+x]=0;
            TestAssert.True(CtDrrProjector.IntegrateRay(near,Front(),20,0,CancellationToken.None)>1.9,
                "Object halfway to the source must project with 2x magnification.");
            Near(0,CtDrrProjector.IntegrateRay(near,Front(),10,0,CancellationToken.None),0);
            near.Origin.Y=-0.5;
            TestAssert.True(CtDrrProjector.IntegrateRay(near,Front(),10,0,CancellationToken.None)>1.9);
            Near(0,CtDrrProjector.IntegrateRay(near,Front(),20,0,CancellationToken.None),0);
        }

        public static void RotatedVolumeAndTranslatedFramePreservePath()
        {
            var volume = Uniform(10,21,21,2,new BeamPoint3D(20,-9,-20),0);
            volume.XAxis=new BeamPoint3D(0,1,0);
            volume.YAxis=new BeamPoint3D(-1,0,0);
            Near(20,CtDrrProjector.IntegrateRay(volume,Front(),0,0,CancellationToken.None),1e-8);
            volume.Origin=new BeamPoint3D(143,447,-98);
            var frame=Front(); frame.Source=new BeamPoint3D(123,-544,-78); frame.Isocenter=new BeamPoint3D(123,456,-78);
            Near(20,CtDrrProjector.IntegrateRay(volume,frame,0,0,CancellationToken.None),1e-8);
        }

        public static void InvalidGeometryAndBudgetsAreRejected()
        {
            Func<CtVolume> valid=()=>Uniform(3,3,3,2,new BeamPoint3D(-2,-2,-2),0);
            var volume=valid(); volume.SizeX=0;
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None));
            volume=valid(); volume.HounsfieldUnits=new float[26];
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None));
            volume=valid(); volume.SpacingZ=double.NaN;
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None));
            volume=valid(); volume.ZAxis=new BeamPoint3D(0,0,-1);
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None));
            volume=valid(); volume.YAxis=new BeamPoint3D(1,0,0);
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None));
            volume=valid(); volume.HounsfieldUnits[4]=float.NaN;
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None));
            volume=valid(); var reversed=Front(); reversed.YAxis=new BeamPoint3D(0,0,-1);
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,reversed,16,20,CancellationToken.None));
            var tilted=Front(); tilted.XAxis=new BeamPoint3D(0,1,0);
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,tilted,16,20,CancellationToken.None));
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),0,20,CancellationToken.None));
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),1025,20,CancellationToken.None));
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,double.PositiveInfinity,CancellationToken.None));
            TestAssert.Throws<ArgumentException>(()=>CtDrrProjector.Project(volume,Front(),16,-1,CancellationToken.None));
        }

        public static void CancellationIsObserved()
        {
            using(var source=new CancellationTokenSource())
            {
                source.Cancel();
                var volume=Uniform(3,3,3,2,new BeamPoint3D(-2,-2,-2),0);
                TestAssert.Throws<OperationCanceledException>(()=>CtDrrProjector.Project(volume,Front(),16,20,source.Token));
                TestAssert.Throws<OperationCanceledException>(()=>CtDrrProjector.IntegrateRay(volume,Front(),0,0,source.Token));
            }
        }

        public static void PixelsAndCtAreNeverSerialized()
        {
            var volume=Uniform(3,3,3,2,new BeamPoint3D(-2,-2,-2),0);
            var image=CtDrrProjector.Project(volume,Front(),16,20,CancellationToken.None);
            string json=JsonConvert.SerializeObject(image);
            TestAssert.False(json.Contains("GrayscalePixels"),"DRR pixels must not enter review JSON.");
            TestAssert.False(JsonConvert.SerializeObject(volume).Contains("HounsfieldUnits"),"CT samples must not enter JSON.");
            TestAssert.True(json.Contains("ProjectionDescription"),"Provenance remains serializable.");
        }

        public static void HfsCardinalFramesAndCollimatorAreExplicit()
        {
            var zero=SyntheticDrrFactory.CreateHfsFrame(0,0,0,1000);
            TestAssert.NotNull(zero,"Synthetic HFS frame is not implemented.");
            Near(-1000,zero.Source.Y,1e-8); Near(1,zero.XAxis.X,1e-8); Near(1,zero.YAxis.Z,1e-8);
            var side=SyntheticDrrFactory.CreateHfsFrame(90,0,0,1000);
            Near(1000,side.Source.X,1e-8); Near(1,side.XAxis.Y,1e-8); Near(1,side.YAxis.Z,1e-8);
            var back=SyntheticDrrFactory.CreateHfsFrame(180,0,0,1000);
            Near(1000,back.Source.Y,1e-8); Near(-1,back.XAxis.X,1e-8);
            var opposite=SyntheticDrrFactory.CreateHfsFrame(270,0,0,1000);
            Near(-1000,opposite.Source.X,1e-8); Near(-1,opposite.XAxis.Y,1e-8);
            var coll=SyntheticDrrFactory.CreateHfsFrame(0,90,0,1000);
            var point=MeshTargetProjector.ProjectPoint(new BeamPoint3D(10,0,20),coll);
            Near(-20,point.X,1e-8); Near(10,point.Y,1e-8);
            TestAssert.Throws<ArgumentException>(()=>SyntheticDrrFactory.CreateHfsFrame(0,0,5,1000));
            TestAssert.Throws<ArgumentException>(()=>SyntheticDrrFactory.CreateHfsFrame(double.NaN,0,0,1000));
        }

        public static void SyntheticSingleAndDualLayerHaveAnatomicalContrast()
        {
            foreach(bool dual in new[] {false,true})
            {
                var beam=SyntheticPlanAnalysisFactory.Create(dual).Beams[0];
                var cp=beam.ControlPoints[0];
                var image=SyntheticDrrFactory.Create(beam,cp);
                TestAssert.NotNull(image,"Synthetic DRR phantom is not implemented.");
                TestAssert.True(image.Synthetic);
                TestAssert.Equal("synthetic",image.SourceStatus);
                TestAssert.Equal(cp.Index,image.ControlPointIndex);
                TestAssert.Equal(cp.GantryAngleDegrees,image.GantryAngleDegrees);
                TestAssert.True(image.GrayscalePixels.Distinct().Count()>32,"Phantom must have bone/lung/body contrast.");
                TestAssert.True(image.GrayscalePixels.Any(p=>p==0));
                var turned=SyntheticDrrFactory.Create(beam,beam.ControlPoints[15]);
                TestAssert.False(image.GrayscalePixels.SequenceEqual(turned.GrayscalePixels),"DRR must follow the selected CP.");
            }
        }

        private static CtVolume Uniform(int x,int y,int z,double spacing,BeamPoint3D origin,float hu)
        {
            return new CtVolume { SizeX=x,SizeY=y,SizeZ=z,SpacingX=spacing,SpacingY=spacing,SpacingZ=spacing,
                Origin=origin,XAxis=new BeamPoint3D(1,0,0),YAxis=new BeamPoint3D(0,1,0),ZAxis=new BeamPoint3D(0,0,1),
                HounsfieldUnits=Enumerable.Repeat(hu,x*y*z).ToArray() };
        }

        private static BeamProjectionFrame Front()
        {
            return new BeamProjectionFrame { Isocenter=new BeamPoint3D(),Source=new BeamPoint3D(0,-1000,0),
                XAxis=new BeamPoint3D(1,0,0),YAxis=new BeamPoint3D(0,0,1) };
        }

        private static void Near(double expected,double actual,double tolerance)
        { TestAssert.True(Math.Abs(expected-actual)<=tolerance,"Expected "+expected+"; got "+actual); }
    }
}
