using System;
using System.Collections.Generic;
using System.Threading;

namespace ClearPlan.Core.PlanAnalysis
{
    public static class CtDrrProjector
    {
        public const int MaximumImageSize = 512;
        public const int MaximumVoxelCount = 256 * 256 * 256;
        private const int MaximumRaySteps = 2048;
        private const long MaximumSampleCount = 100000000;
        private const double AxisTolerance = 1e-6;
        public const string MethodDescription =
            "Divergent source-to-isocenter-plane DRR; finite voxel support and trilinear HU sampling. " +
            "Integral of max(0, 1 + HU/1000) over path length in mm, with positive-integral percentile windowing. " +
            "Attenuation proxy only: not dose, calibrated electron density, diagnostic imaging or registration-grade imaging.";

        /// <summary>Square overview image. Row zero is positive beam Y, columns increase along beam X.</summary>
        public static BeamEyeViewImage Project(CtVolume volume,BeamProjectionFrame frame,int size,double halfExtentMm,CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if(size<1 || size>MaximumImageSize || !Finite(halfExtentMm) || halfExtentMm<=0 || halfExtentMm>1000)
                throw new ArgumentException("DRR extent or resolution exceeds the bounded overview raster.");
            var context=Validate(volume,frame,token);
            var integrals=new double[size*size];
            long remainingSamples=MaximumSampleCount;
            double pixelSpacing=2*halfExtentMm/size;
            for(int y=0;y<size;y++)
            {
                token.ThrowIfCancellationRequested();
                for(int x=0;x<size;x++)
                    integrals[y*size+x]=Integrate(context,-halfExtentMm+(x+0.5)*pixelSpacing,
                        halfExtentMm-(y+0.5)*pixelSpacing,token,ref remainingSamples);
            }
            return new BeamEyeViewImage { WidthPixels=size,HeightPixels=size,ExtentMm=halfExtentMm,
                GrayscalePixels=Window(integrals,token),SourceStatus="available",ProjectionDescription=MethodDescription };
        }

        /// <summary>Unwindowed HU attenuation-proxy integral; not a dose or physical electron-density calculation.</summary>
        public static double IntegrateRay(CtVolume volume,BeamProjectionFrame frame,double xMm,double yMm,CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if(!Finite(xMm) || !Finite(yMm) || Math.Abs(xMm)>1000 || Math.Abs(yMm)>1000)
                throw new ArgumentException("Invalid isocenter-plane ray coordinate.");
            var context=Validate(volume,frame,token);
            long remainingSamples=MaximumSampleCount;
            return Integrate(context,xMm,yMm,token,ref remainingSamples);
        }

        private static double Integrate(Context context,double xMm,double yMm,CancellationToken token,ref long remainingSamples)
        {
            var direction=(context.Isocenter+context.FrameX*xMm+context.FrameY*yMm-context.Source).Unit();
            var localDirection=context.ToVoxelDirection(direction);
            double entry=0,exit=double.PositiveInfinity;
            var volume=context.Volume;
            if(!Clip(context.LocalSource.X,localDirection.X,-0.5,volume.SizeX-0.5,ref entry,ref exit) ||
                !Clip(context.LocalSource.Y,localDirection.Y,-0.5,volume.SizeY-0.5,ref entry,ref exit) ||
                !Clip(context.LocalSource.Z,localDirection.Z,-0.5,volume.SizeZ-0.5,ref entry,ref exit) || exit<=entry)
                return 0;
            double length=exit-entry;
            int steps=(int)Math.Ceiling(length/context.StepMm);
            if(steps<1 || steps>MaximumRaySteps || remainingSamples<steps)
                throw new ArgumentException("CT ray integration exceeds the bounded overview sampling budget.");
            remainingSamples-=steps;
            double step=length/steps,integral=0;
            var increment=localDirection*step;
            var voxel=context.LocalSource+localDirection*(entry+0.5*step);
            for(int i=0;i<steps;i++,voxel=voxel+increment)
            {
                if((i&127)==0) token.ThrowIfCancellationRequested();
                integral+=Math.Max(0,1+Sample(volume,voxel)/1000.0)*step;
            }
            return integral;
        }

        private static bool Clip(double origin,double direction,double lower,double upper,ref double entry,ref double exit)
        {
            if(Math.Abs(direction)<1e-12) return origin>=lower && origin<=upper;
            double a=(lower-origin)/direction,b=(upper-origin)/direction;
            if(a>b) { double swap=a; a=b; b=swap; }
            entry=Math.Max(entry,a); exit=Math.Min(exit,b);
            return exit>entry;
        }

        // The half-voxel boundary region uses the nearest edge value. Outside the exact support is air.
        private static double Sample(CtVolume volume,Vector voxel)
        {
            double x=Math.Max(0,Math.Min(volume.SizeX-1,voxel.X));
            double y=Math.Max(0,Math.Min(volume.SizeY-1,voxel.Y));
            double z=Math.Max(0,Math.Min(volume.SizeZ-1,voxel.Z));
            int x0=(int)x,y0=(int)y,z0=(int)z;
            int x1=Math.Min(volume.SizeX-1,x0+1),y1=Math.Min(volume.SizeY-1,y0+1),z1=Math.Min(volume.SizeZ-1,z0+1);
            double dx=x-x0,dy=y-y0,dz=z-z0;
            int row00=(z0*volume.SizeY+y0)*volume.SizeX,row01=(z0*volume.SizeY+y1)*volume.SizeX;
            int row10=(z1*volume.SizeY+y0)*volume.SizeX,row11=(z1*volume.SizeY+y1)*volume.SizeX;
            var hu=volume.HounsfieldUnits;
            double a=Lerp(hu[row00+x0],hu[row00+x1],dx),b=Lerp(hu[row01+x0],hu[row01+x1],dx);
            double c=Lerp(hu[row10+x0],hu[row10+x1],dx),d=Lerp(hu[row11+x0],hu[row11+x1],dx);
            return Lerp(Lerp(a,b,dy),Lerp(c,d,dy),dz);
        }

        private static byte[] Window(double[] values,CancellationToken token)
        {
            var positive=new List<double>();
            foreach(double value in values) if(value>0) positive.Add(value);
            var pixels=new byte[values.Length];
            if(positive.Count==0) return pixels;
            positive.Sort();
            token.ThrowIfCancellationRequested();
            double low=positive[(int)((positive.Count-1)*0.01)],high=positive[(int)((positive.Count-1)*0.995)];
            if(high-low<Math.Max(1e-9,high*1e-6)) low=0;
            double width=high-low;
            for(int i=0;i<values.Length;i++)
            {
                if((i&4095)==0) token.ThrowIfCancellationRequested();
                pixels[i]=(byte)Math.Round(255*Math.Max(0,Math.Min(1,(values[i]-low)/width)));
            }
            return pixels;
        }

        private static Context Validate(CtVolume volume,BeamProjectionFrame frame,CancellationToken token)
        {
            if(volume==null) throw new ArgumentNullException("volume");
            if(frame==null) throw new ArgumentNullException("frame");
            if(volume.SizeX<1 || volume.SizeY<1 || volume.SizeZ<1 ||
                volume.SizeX>512 || volume.SizeY>512 || volume.SizeZ>512 ||
                (long)volume.SizeX*volume.SizeY*volume.SizeZ>MaximumVoxelCount || volume.HounsfieldUnits==null ||
                volume.HounsfieldUnits.Length!=(long)volume.SizeX*volume.SizeY*volume.SizeZ)
                throw new ArgumentException("CT dimensions or sample count exceed the bounded volume.");
            RequireSpacing(volume.SpacingX,volume.SizeX); RequireSpacing(volume.SpacingY,volume.SizeY); RequireSpacing(volume.SpacingZ,volume.SizeZ);
            var vx=Vector.From(volume.XAxis);var vy=Vector.From(volume.YAxis);var vz=Vector.From(volume.ZAxis);
            RequireBasis(vx,vy,vz,"CT axes must be a right-handed orthonormal patient-space basis.");
            var source=Vector.From(frame.Source);var iso=Vector.From(frame.Isocenter);
            double sad=(source-iso).Length;
            if(sad<100 || sad>5000) throw new ArgumentException("Source-to-isocenter distance is outside the supported overview range.");
            var fx=Vector.From(frame.XAxis);var fy=Vector.From(frame.YAxis);
            RequireBasis(fx,fy,(source-iso).Unit(),"Beam X cross Y must point toward the source in an orthonormal frame.");
            var origin=Vector.From(volume.Origin);
            for(int i=0;i<volume.HounsfieldUnits.Length;i++)
            {
                if((i&16383)==0) token.ThrowIfCancellationRequested();
                if(!Finite(volume.HounsfieldUnits[i])) throw new ArgumentException("CT samples contain non-finite HU values.");
            }
            return new Context { Volume=volume,VolumeX=vx,VolumeY=vy,VolumeZ=vz,
                Source=source,Isocenter=iso,FrameX=fx,FrameY=fy,
                LocalSource=new Vector(Vector.Dot(source-origin,vx)/volume.SpacingX,
                    Vector.Dot(source-origin,vy)/volume.SpacingY,Vector.Dot(source-origin,vz)/volume.SpacingZ),
                StepMm=0.5*Math.Min(volume.SpacingX,Math.Min(volume.SpacingY,volume.SpacingZ)) };
        }

        private static void RequireSpacing(double spacing,int size)
        {
            if(!Finite(spacing) || spacing<0.05 || spacing>100 || spacing*size>5000)
                throw new ArgumentException("CT spacing or physical extent is outside the bounded overview range.");
        }

        private static void RequireBasis(Vector x,Vector y,Vector z,string reason)
        {
            if(Math.Abs(x.Length-1)>AxisTolerance || Math.Abs(y.Length-1)>AxisTolerance || Math.Abs(z.Length-1)>AxisTolerance ||
                Math.Abs(Vector.Dot(x,y))>AxisTolerance || Math.Abs(Vector.Dot(x,z))>AxisTolerance ||
                Math.Abs(Vector.Dot(y,z))>AxisTolerance || Vector.Dot(Vector.Cross(x,y),z)<1-AxisTolerance)
                throw new ArgumentException(reason);
        }

        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static double Lerp(double a,double b,double fraction) { return a+(b-a)*fraction; }

        private sealed class Context
        {
            internal CtVolume Volume;
            internal Vector VolumeX,VolumeY,VolumeZ,Source,Isocenter,FrameX,FrameY,LocalSource;
            internal double StepMm;
            internal Vector ToVoxelDirection(Vector point)
            { return new Vector(Vector.Dot(point,VolumeX)/Volume.SpacingX,Vector.Dot(point,VolumeY)/Volume.SpacingY,Vector.Dot(point,VolumeZ)/Volume.SpacingZ); }
        }

        private struct Vector
        {
            internal double X,Y,Z;
            internal Vector(double x,double y,double z) { X=x;Y=y;Z=z; }
            internal double Length { get { return Math.Sqrt(Dot(this,this)); } }
            internal Vector Unit() { return this*(1/Length); }
            internal static Vector From(BeamPoint3D point)
            {
                if(point==null || !Finite(point.X) || !Finite(point.Y) || !Finite(point.Z) ||
                    Math.Abs(point.X)>100000 || Math.Abs(point.Y)>100000 || Math.Abs(point.Z)>100000)
                    throw new ArgumentException("Missing, non-finite or out-of-range spatial coordinate.");
                return new Vector(point.X,point.Y,point.Z);
            }
            public static Vector operator +(Vector a,Vector b) { return new Vector(a.X+b.X,a.Y+b.Y,a.Z+b.Z); }
            public static Vector operator -(Vector a,Vector b) { return new Vector(a.X-b.X,a.Y-b.Y,a.Z-b.Z); }
            public static Vector operator *(Vector a,double scale) { return new Vector(a.X*scale,a.Y*scale,a.Z*scale); }
            internal static double Dot(Vector a,Vector b) { return a.X*b.X+a.Y*b.Y+a.Z*b.Z; }
            internal static Vector Cross(Vector a,Vector b) { return new Vector(a.Y*b.Z-a.Z*b.Y,a.Z*b.X-a.X*b.Z,a.X*b.Y-a.Y*b.X); }
        }
    }
}
