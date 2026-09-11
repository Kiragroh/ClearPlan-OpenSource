using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace ClearPlan.Core.PlanAnalysis
{
    public sealed class BeamPoint3D
    {
        public BeamPoint3D() { }
        public BeamPoint3D(double x,double y,double z) { X=x; Y=y; Z=z; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }
    public sealed class TargetMeshGeometry
    {
        public List<BeamPoint3D> Vertices { get; set; }
        public int[] TriangleIndices { get; set; }
    }
    public sealed class BeamProjectionFrame
    {
        public BeamPoint3D Isocenter { get; set; }
        public BeamPoint3D Source { get; set; }
        public BeamPoint3D XAxis { get; set; }
        public BeamPoint3D YAxis { get; set; }
    }
    public static class MeshTargetProjector
    {
        public const double DefaultResolutionMm = 0.625;
        private const int MaximumRasterDimension = 4096;

        public static double NativeCoordinateRotation(List<List<BeamPoint>> zeroCollimator,
            List<List<BeamPoint>> beamCoordinates, double nativeCollimatorDegrees)
        {
            // ESAPI re-rasterizes these silhouettes independently: point counts/order need not match.
            // The native collimator magnitude supplies only candidates; the actual outline geometry
            // must uniquely establish the sign. No fitted angle, IEC sign or replacement frame is guessed.
            ValidateNativeOutlines(zeroCollimator); ValidateNativeOutlines(beamCoordinates);
            if(!PlanAnalysisCalculator.Finite(nativeCollimatorDegrees)) throw new ArgumentException("Invalid native collimator angle.");
            long firstCount=zeroCollimator.Sum(l=>(long)l.Count),secondCount=beamCoordinates.Sum(l=>(long)l.Count);
            if(firstCount*secondCount>16000000) throw new ArgumentException("Native outline calibration exceeds its bounded comparison budget.");
            double angle=Math.IEEERemainder(nativeCollimatorDegrees,360)*Math.PI/180;
            var candidates=new List<double> { angle };
            if(Math.Abs(Math.IEEERemainder(2*angle,2*Math.PI))>1e-8) candidates.Add(-angle);
            var matches=new List<double>();
            foreach(double candidate in candidates)
            {
                double c=Math.Cos(candidate),s=Math.Sin(candidate);
                var rotated=zeroCollimator.Select(l=>l.Select(p=>new BeamPoint(c*p.X-s*p.Y,s*p.X+c*p.Y)).ToList()).ToList();
                // Absolute geometric tolerance in the isocenter plane, not a clinical acceptance limit.
                if(BoundariesWithin(rotated,beamCoordinates,0.5) && BoundariesWithin(beamCoordinates,rotated,0.5)) matches.Add(candidate);
            }
            if(matches.Count!=1) throw new ArgumentException("Native silhouette rotation is inconsistent or ambiguous within 0.5 mm.");
            return matches[0];
        }

        private static void ValidateNativeOutlines(List<List<BeamPoint>> outlines)
        {
            if(outlines==null || outlines.Count==0 || outlines.Count>512 ||
                outlines.Any(l=>l==null || l.Count<3) || outlines.Sum(l=>(long)l.Count)>16000 ||
                outlines.SelectMany(l=>l).Any(p=>p==null || !PlanAnalysisCalculator.Finite(p.X) ||
                    !PlanAnalysisCalculator.Finite(p.Y) || Math.Abs(p.X)>2000 || Math.Abs(p.Y)>2000))
                throw new ArgumentException("Invalid native outlines.");
        }

        private static bool BoundariesWithin(List<List<BeamPoint>> source,List<List<BeamPoint>> destination,double tolerance)
        {
            foreach(var loop in source) for(int i=0;i<loop.Count;i++)
            {
                var a=loop[i]; var b=loop[(i+1)%loop.Count];
                if(!NearBoundary(a,destination,tolerance*tolerance) ||
                    !NearBoundary(new BeamPoint((a.X+b.X)/2,(a.Y+b.Y)/2),destination,tolerance*tolerance)) return false;
            }
            return true;
        }

        private static bool NearBoundary(BeamPoint p,List<List<BeamPoint>> outlines,double squaredTolerance)
        {
            foreach(var loop in outlines) for(int i=0;i<loop.Count;i++)
            {
                var a=loop[i];var b=loop[(i+1)%loop.Count];
                double dx=b.X-a.X,dy=b.Y-a.Y,denominator=dx*dx+dy*dy;
                double t=denominator>1e-20 ? Math.Max(0,Math.Min(1,((p.X-a.X)*dx+(p.Y-a.Y)*dy)/denominator)) : 0;
                double x=p.X-a.X-t*dx,y=p.Y-a.Y-t*dy;
                if(x*x+y*y<=squaredTolerance) return true;
            }
            return false;
        }

        public static double NativeCoordinateRotation(List<List<BeamPoint>> zeroCollimator,List<List<BeamPoint>> beamCoordinates)
        {
            if (zeroCollimator==null || beamCoordinates==null || zeroCollimator.Count!=beamCoordinates.Count)
                throw new ArgumentException("Native outline correspondence is unavailable.");
            for(int i=0;i<zeroCollimator.Count;i++)
                if(zeroCollimator[i]==null || beamCoordinates[i]==null || zeroCollimator[i].Count!=beamCoordinates[i].Count)
                    throw new ArgumentException("Native outline correspondence is unavailable.");
            var first=zeroCollimator.SelectMany(l=>l).ToList();
            var second=beamCoordinates.SelectMany(l=>l).ToList();
            if(first.Count<3) throw new ArgumentException("Native outlines are empty.");
            double dot=0,cross=0;
            for(int i=0;i<first.Count;i++)
            { dot+=first[i].X*second[i].X+first[i].Y*second[i].Y; cross+=first[i].X*second[i].Y-first[i].Y*second[i].X; }
            if(!PlanAnalysisCalculator.Finite(dot)||!PlanAnalysisCalculator.Finite(cross)||Math.Abs(dot)+Math.Abs(cross)<1e-9)
                throw new ArgumentException("Native coordinate rotation cannot be determined.");
            double angle=Math.Atan2(cross,dot),c=Math.Cos(angle),s=Math.Sin(angle);
            for(int i=0;i<first.Count;i++)
                if(Math.Abs(c*first[i].X-s*first[i].Y-second[i].X)>0.02 ||
                    Math.Abs(s*first[i].X+c*first[i].Y-second[i].Y)>0.02)
                    throw new ArgumentException("Native outlines do not correspond by a single origin-centered rotation.");
            return angle;
        }

        public static BeamPoint ProjectPoint(BeamPoint3D point, BeamProjectionFrame frame)
        {
            ValidateFrame(frame);
            return ProjectUnchecked(point,frame);
        }

        // The angle is a 2D coordinate rotation measured from paired native BEV/beam outlines,
        // not a guessed IEC collimator sign. The source positions determine the gantry plane.
        public static BeamProjectionFrame CreateFrame(BeamPoint3D iso, BeamPoint3D source,
            BeamPoint3D sourceAtZero, BeamPoint3D sourceAtNinety, double rotationRadians)
        {
            RequireFinite(iso); RequireFinite(source); RequireFinite(sourceAtZero); RequireFinite(sourceAtNinety);
            if (!PlanAnalysisCalculator.Finite(rotationRadians)) throw new ArgumentException("Invalid coordinate rotation.");
            var n = Unit(Subtract(source,iso));
            var v = Unit(Cross(Unit(Subtract(sourceAtZero,iso)),Unit(Subtract(sourceAtNinety,iso))));
            if (Math.Abs(Dot(n,v))>1e-5) throw new ArgumentException("Source is outside the vendor gantry plane.");
            var u = Unit(Cross(v,n));
            double c = Math.Cos(rotationRadians), s = Math.Sin(rotationRadians);
            var frame = new BeamProjectionFrame { Isocenter=iso, Source=source,
                XAxis=Add(Scale(u,c),Scale(v,-s)), YAxis=Add(Scale(u,s),Scale(v,c)) };
            ValidateFrame(frame);
            return frame;
        }

        public static List<ApertureRectangle> Project(TargetMeshGeometry mesh, BeamProjectionFrame frame, double resolutionMm)
        { return Project(mesh,frame,resolutionMm,CancellationToken.None); }

        public static List<ApertureRectangle> Project(TargetMeshGeometry mesh, BeamProjectionFrame frame, double resolutionMm,
            CancellationToken cancellationToken)
        {
            ValidateResolution(resolutionMm);
            ValidateFrame(frame);
            if (mesh == null || mesh.Vertices == null || mesh.TriangleIndices == null ||
                mesh.Vertices.Count < 3 || mesh.TriangleIndices.Length == 0 || mesh.TriangleIndices.Length%3 != 0 ||
                mesh.TriangleIndices.Any(i => i<0 || i>=mesh.Vertices.Count) || mesh.TriangleIndices.Length>750000)
                throw new ArgumentException("Incomplete triangulated target surface.");
            cancellationToken.ThrowIfCancellationRequested();
            var points = mesh.Vertices.Select(p => ProjectUnchecked(p,frame)).ToArray();
            ValidateExtent(points,resolutionMm);
            var rows = new Dictionary<int,List<Span>>();
            int spanBudget=4000000;
            for (int i=0;i<mesh.TriangleIndices.Length;i+=3)
            {
                if(i%768==0) cancellationToken.ThrowIfCancellationRequested();
                var triangle = new[] { points[mesh.TriangleIndices[i]],points[mesh.TriangleIndices[i+1]],points[mesh.TriangleIndices[i+2]] };
                int first = RowAt(triangle.Min(p=>p.Y),resolutionMm);
                int stop = RowAt(triangle.Max(p=>p.Y),resolutionMm);
                for (int row=first;row<stop;row++)
                {
                    if(--spanBudget<0) throw new ArgumentException("Projected triangle spans exceed the analysis budget.");
                    var xs = Crossings(triangle,(row+0.5)*resolutionMm);
                    if (xs.Count < 2) continue; // Edge-on projected surface triangle.
                    AddSpan(rows,row,xs.Min(),xs.Max(),resolutionMm);
                }
            }
            return MergeRows(rows,resolutionMm);
        }

        public static List<ApertureRectangle> RasterizeOutlines(List<List<BeamPoint>> outlines, double resolutionMm)
        {
            ValidateResolution(resolutionMm);
            if (outlines == null || outlines.Count==0 || outlines.Any(l=>l==null || l.Count<3))
                throw new ArgumentException("Missing native target outlines.");
            var points = outlines.SelectMany(l=>l).ToArray();
            ValidateExtent(points,resolutionMm);
            var rows = new Dictionary<int,List<Span>>();
            int first = RowAt(points.Min(p=>p.Y),resolutionMm), stop = RowAt(points.Max(p=>p.Y),resolutionMm);
            for (int row=first;row<stop;row++)
            {
                var xs = outlines.SelectMany(l=>Crossings(l,(row+0.5)*resolutionMm)).OrderBy(x=>x).ToList();
                if (xs.Count%2 != 0) throw new ArgumentException("Invalid native contour topology.");
                for (int i=0;i<xs.Count;i+=2) AddSpan(rows,row,xs[i],xs[i+1],resolutionMm);
            }
            return MergeRows(rows,resolutionMm);
        }

        // Both arguments are unions of non-overlapping horizontal rectangles.
        public static double SymmetricDifferenceFraction(List<ApertureRectangle> a,List<ApertureRectangle> b)
        {
            double areaA=AreaMm2(a),areaB=AreaMm2(b),intersection=IntersectionAreaMm2(a,b);
            double union=areaA+areaB-intersection;
            return union>0 ? Math.Max(0,(areaA+areaB-2*intersection)/union) : 1;
        }

        public static double AreaMm2(List<ApertureRectangle> rows)
        { return rows.Sum(r=>(r.X2-r.X1)*(r.Y2-r.Y1)); }

        public static double IntersectionAreaMm2(List<ApertureRectangle> a,List<ApertureRectangle> b)
        {
            double area=0;
            foreach (var first in a) foreach (var second in b)
            {
                double h=Math.Min(first.Y2,second.Y2)-Math.Max(first.Y1,second.Y1);
                if (h<=0) continue;
                area += h*Math.Max(0,Math.Min(first.X2,second.X2)-Math.Max(first.X1,second.X1));
            }
            return area;
        }

        private static BeamPoint ProjectUnchecked(BeamPoint3D point,BeamProjectionFrame frame)
        {
            RequireFinite(point);
            var source=Subtract(frame.Source,frame.Isocenter);
            double sad=Length(source);
            var relative=Subtract(point,frame.Isocenter);
            double distance=sad-Dot(relative,Scale(source,1/sad));
            if (distance<=1e-6) throw new ArgumentException("Target vertex is at or behind the radiation source.");
            double magnification=sad/distance;
            return new BeamPoint(magnification*Dot(relative,frame.XAxis),magnification*Dot(relative,frame.YAxis));
        }

        private static List<double> Crossings(IList<BeamPoint> polygon,double y)
        {
            var xs=new List<double>();
            for (int i=0;i<polygon.Count;i++)
            {
                var a=polygon[i]; var b=polygon[(i+1)%polygon.Count];
                if ((a.Y<=y && b.Y>y)||(b.Y<=y && a.Y>y)) xs.Add(a.X+(y-a.Y)*(b.X-a.X)/(b.Y-a.Y));
            }
            return xs;
        }

        private static int RowAt(double coordinate,double resolution)
        { return checked((int)Math.Ceiling(coordinate/resolution-0.5)); }

        private static void AddSpan(Dictionary<int,List<Span>> rows,int row,double min,double max,double resolution)
        {
            int first=RowAt(min,resolution),stop=RowAt(max,resolution);
            if (stop<=first) return;
            List<Span> spans;
            if (!rows.TryGetValue(row,out spans)) { spans=new List<Span>(); rows.Add(row,spans); }
            spans.Add(new Span { First=first,Stop=stop });
        }

        private static List<ApertureRectangle> MergeRows(Dictionary<int,List<Span>> rows,double resolution)
        {
            var result=new List<ApertureRectangle>();
            foreach (var row in rows.OrderBy(r=>r.Key))
            {
                var spans=row.Value.OrderBy(s=>s.First).ThenBy(s=>s.Stop).ToList();
                int first=spans[0].First,stop=spans[0].Stop;
                for (int i=1;i<spans.Count;i++)
                {
                    if (spans[i].First<=stop) stop=Math.Max(stop,spans[i].Stop);
                    else { result.Add(new ApertureRectangle(first*resolution,row.Key*resolution,stop*resolution,(row.Key+1)*resolution)); first=spans[i].First;stop=spans[i].Stop; }
                }
                result.Add(new ApertureRectangle(first*resolution,row.Key*resolution,stop*resolution,(row.Key+1)*resolution));
            }
            return result;
        }

        private static void ValidateExtent(BeamPoint[] points,double resolution)
        {
            if (points.Length<3 || points.Any(p=>p==null || !PlanAnalysisCalculator.Finite(p.X) || !PlanAnalysisCalculator.Finite(p.Y)))
                throw new ArgumentException("Invalid projected target points.");
            if ((points.Max(p=>p.X)-points.Min(p=>p.X))/resolution>MaximumRasterDimension ||
                (points.Max(p=>p.Y)-points.Min(p=>p.Y))/resolution>MaximumRasterDimension ||
                points.Any(p=>Math.Abs(p.X/resolution)>int.MaxValue/2 || Math.Abs(p.Y/resolution)>int.MaxValue/2))
                throw new ArgumentException("Projected target exceeds bounded analysis raster.");
        }
        private static void ValidateResolution(double resolution)
        { if (!PlanAnalysisCalculator.Finite(resolution)||resolution<=0) throw new ArgumentException("Invalid raster resolution."); }
        private static void ValidateFrame(BeamProjectionFrame frame)
        {
            if (frame==null) throw new ArgumentNullException("frame");
            RequireFinite(frame.Source); RequireFinite(frame.Isocenter); RequireFinite(frame.XAxis);RequireFinite(frame.YAxis);
            var n=Unit(Subtract(frame.Source,frame.Isocenter));
            if (Math.Abs(Length(frame.XAxis)-1)>1e-6 || Math.Abs(Length(frame.YAxis)-1)>1e-6 ||
                Math.Abs(Dot(frame.XAxis,frame.YAxis))>1e-6 || Math.Abs(Dot(n,frame.XAxis))>1e-6 || Math.Abs(Dot(n,frame.YAxis))>1e-6)
                throw new ArgumentException("Projection frame is not orthonormal on the isocenter plane.");
        }
        private static void RequireFinite(BeamPoint3D p)
        { if(p==null || !PlanAnalysisCalculator.Finite(p.X)||!PlanAnalysisCalculator.Finite(p.Y)||!PlanAnalysisCalculator.Finite(p.Z)) throw new ArgumentException("Invalid 3D point."); }
        private static BeamPoint3D Subtract(BeamPoint3D a,BeamPoint3D b) { return new BeamPoint3D(a.X-b.X,a.Y-b.Y,a.Z-b.Z); }
        private static BeamPoint3D Add(BeamPoint3D a,BeamPoint3D b) { return new BeamPoint3D(a.X+b.X,a.Y+b.Y,a.Z+b.Z); }
        private static BeamPoint3D Scale(BeamPoint3D p,double s) { return new BeamPoint3D(p.X*s,p.Y*s,p.Z*s); }
        private static double Dot(BeamPoint3D a,BeamPoint3D b) { return a.X*b.X+a.Y*b.Y+a.Z*b.Z; }
        private static double Length(BeamPoint3D p) { return Math.Sqrt(Dot(p,p)); }
        private static BeamPoint3D Unit(BeamPoint3D p) { double l=Length(p);if(l<1e-9) throw new ArgumentException("Degenerate source frame.");return Scale(p,1/l); }
        private static BeamPoint3D Cross(BeamPoint3D a,BeamPoint3D b) { return new BeamPoint3D(a.Y*b.Z-a.Z*b.Y,a.Z*b.X-a.X*b.Z,a.X*b.Y-a.Y*b.X); }
        private sealed class Span { internal int First,Stop; }
    }
}
