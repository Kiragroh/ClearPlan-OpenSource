using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    /// <summary>Analytic point-to-solid distances and 1-Lipschitz bounds over supplied triangles, not vertex-only minima.</summary>
    internal static class CollisionSweepDistance
    {
        internal sealed class PreparedSurface
        {
            public CollisionSurface Surface;
            public bool Valid, Closed, Degenerate, RadialSurfaceValid, HasSurfaceFaces;
            public string Reason;
            public SpatialNode SpatialRoot;
            public int[] SpatialTriangles;
        }
        internal sealed class SpatialNode
        {
            public BeamPoint3D Center;
            public double Radius;
            public int Start, Count;
            public SpatialNode Left, Right;
        }
        internal sealed class DistanceBudget
        {
            private long remaining = CollisionSweep.MaximumTrianglePosePairs;
            public void Charge(long count) { remaining -= count; if (remaining < 0) throw new ArgumentException("Source distance evaluation exhausted its bounded spatial workload."); }
        }
        private sealed class Edge { public int Count, Orientation; }

        internal static PreparedSurface Prepare(CollisionSurface surface, CancellationToken token)
        {
            var result = new PreparedSurface { Surface = surface, Reason = surface.Role + " mesh is missing or invalid." };
            var mesh = surface.Mesh;
            if (mesh == null || mesh.Vertices == null || mesh.TriangleIndices == null || mesh.Vertices.Count < 3 ||
                mesh.Vertices.Count > 250000 || mesh.TriangleIndices.Length < 3 || mesh.TriangleIndices.Length > 1500000 || mesh.TriangleIndices.Length % 3 != 0) return result;
            for (int i = 0; i < mesh.Vertices.Count; i++)
            { if ((i & 255) == 0) token.ThrowIfCancellationRequested(); if (!CollisionSweep.ValidPoint(mesh.Vertices[i])) return result; }
            var edges = new Dictionary<long, Edge>(); bool degenerate = false;
            var properVertex = new bool[mesh.Vertices.Count];
            var referencedVertex = new bool[mesh.Vertices.Count];
            for (int i = 0; i < mesh.TriangleIndices.Length; i += 3)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                int a = mesh.TriangleIndices[i], b = mesh.TriangleIndices[i + 1], c = mesh.TriangleIndices[i + 2];
                if (a < 0 || b < 0 || c < 0 || a >= mesh.Vertices.Count || b >= mesh.Vertices.Count || c >= mesh.Vertices.Count) return result;
                var av = mesh.Vertices[a]; var bv = mesh.Vertices[b]; var cv = mesh.Vertices[c];
                double ux = bv.X - av.X, uy = bv.Y - av.Y, uz = bv.Z - av.Z;
                double vx = cv.X - av.X, vy = cv.Y - av.Y, vz = cv.Z - av.Z;
                double nx = uy * vz - uz * vy, ny = uz * vx - ux * vz, nz = ux * vy - uy * vx;
                double longest = Math.Max(Distance2(av, bv), Math.Max(Distance2(av, cv), Distance2(bv, cv)));
                bool zeroArea = nx * nx + ny * ny + nz * nz <= Math.Max(1e-16, longest * longest * 1e-12);
                degenerate |= zeroArea;
                referencedVertex[a] = referencedVertex[b] = referencedVertex[c] = true;
                if (!zeroArea) properVertex[a] = properVertex[b] = properVertex[c] = true;
                AddEdge(edges, a, b); AddEdge(edges, b, c); AddEdge(edges, c, a);
            }
            result.Valid = true; result.Degenerate = degenerate; result.Closed = !degenerate && edges.Values.All(e => e.Count == 2 && e.Orientation == 0);
            // WPF can repeat indices/coordinates in zero-area seam triangles. These
            // do not change a radial maximum if all their points belong to actual
            // nondegenerate faces. No tolerance-based welding or surface repair.
            var properPoints = new HashSet<Tuple<double, double, double>>();
            for (int i = 0; i < mesh.Vertices.Count; i++)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                var p = mesh.Vertices[i];
                if (properVertex[i]) properPoints.Add(Tuple.Create(p.X, p.Y, p.Z));
            }
            result.HasSurfaceFaces = properPoints.Count > 0;
            result.RadialSurfaceValid = result.HasSurfaceFaces;
            for (int i = 0; i < mesh.Vertices.Count && result.RadialSurfaceValid; i++)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                var p = mesh.Vertices[i];
                if (referencedVertex[i] && !properVertex[i] && !properPoints.Contains(Tuple.Create(p.X, p.Y, p.Z))) result.RadialSurfaceValid = false;
            }
            result.Reason = result.Closed ? null : surface.Role + " mesh is open, nonmanifold or degenerate.";
            return result;
        }

        private static void AddEdge(Dictionary<long, Edge> edges, int a, int b)
        {
            long key = ((long)Math.Min(a, b) << 32) | (uint)Math.Max(a, b); Edge edge;
            if (!edges.TryGetValue(key, out edge)) { edge = new Edge(); edges.Add(key, edge); }
            edge.Count++; edge.Orientation += a < b ? 1 : -1;
        }

        // Axis-aligned enclosing spheres provide 1-Lipschitz lower bounds. Pruning
        // never removes a possible minimum: its lower bound must exceed an upper
        // bound sampled ON a supplied triangle. No mesh decimation or topology repair.
        internal static void SampleSpatial(PreparedSurface surface, Func<BeamPoint3D, string, double> distance,
            CancellationToken token, DistanceBudget budget, out double? lower, out double? upper)
        {
            var mesh = surface.Surface.Mesh;
            if (mesh.TriangleIndices.Length < 3072)
            {
                budget.Charge(mesh.Vertices.Count + mesh.TriangleIndices.Length / 3);
                Sample(surface, distance, token, out lower, out upper); return;
            }
            if (surface.SpatialRoot == null)
            {
                var unique = new HashSet<Tuple<int,int,int>>();
                var triangles = new List<int>();
                for (int i=0; i<mesh.TriangleIndices.Length; i+=3)
                {
                    if ((i & 255)==0) token.ThrowIfCancellationRequested();
                    int a=mesh.TriangleIndices[i],b=mesh.TriangleIndices[i+1],c=mesh.TriangleIndices[i+2];
                    int lo=Math.Min(a,Math.Min(b,c)), hi=Math.Max(a,Math.Max(b,c));
                    if (unique.Add(Tuple.Create(lo,a+b+c-lo-hi,hi))) triangles.Add(i/3);
                }
                // Repeated index triples describe the same distance domain. Original
                // mesh and its topology/degeneracy classification remain untouched.
                surface.SpatialTriangles = triangles.ToArray();
                surface.SpatialRoot = BuildSpatial(mesh, surface.SpatialTriangles, 0, surface.SpatialTriangles.Length, token);
            }
            budget.Charge(1);
            double minimum = distance(mesh.Vertices[mesh.TriangleIndices[0]], surface.Surface.Role);
            double bound = double.PositiveInfinity;
            Action<SpatialNode, double> visit = null;
            Func<SpatialNode, double> nodeBound = node => { budget.Charge(1); return distance(node.Center, surface.Surface.Role) - node.Radius - 1e-8; };
            visit = (node, nodeLower) =>
            {
                token.ThrowIfCancellationRequested();
                if (nodeLower >= minimum) { bound = Math.Min(bound, nodeLower); return; }
                if (node.Left != null)
                {
                    double l = nodeBound(node.Left), r = nodeBound(node.Right);
                    if (l <= r) { visit(node.Left, l); visit(node.Right, r); }
                    else { visit(node.Right, r); visit(node.Left, l); }
                    return;
                }
                for (int n = node.Start; n < node.Start + node.Count; n++)
                {
                    budget.Charge(4);
                    int i = surface.SpatialTriangles[n] * 3;
                    var a = mesh.Vertices[mesh.TriangleIndices[i]]; var b = mesh.Vertices[mesh.TriangleIndices[i + 1]]; var c = mesh.Vertices[mesh.TriangleIndices[i + 2]];
                    var center = new BeamPoint3D((a.X+b.X+c.X)/3, (a.Y+b.Y+c.Y)/3, (a.Z+b.Z+c.Z)/3);
                    double av = distance(a, surface.Surface.Role), bv = distance(b, surface.Surface.Role), cv = distance(c, surface.Surface.Role), value = distance(center, surface.Surface.Role);
                    minimum = Math.Min(minimum, Math.Min(value, Math.Min(av, Math.Min(bv, cv))));
                    double radius = Math.Sqrt(Math.Max(Distance2(center, a), Math.Max(Distance2(center, b), Distance2(center, c))));
                    double ab = Math.Sqrt(Distance2(a,b)), ac = Math.Sqrt(Distance2(a,c)), bc = Math.Sqrt(Distance2(b,c));
                    bound = Math.Min(bound, Math.Max(value-radius, Math.Max(av-Math.Max(ab,ac), Math.Max(bv-Math.Max(ab,bc), cv-Math.Max(ac,bc)))));
                }
            };
            visit(surface.SpatialRoot, nodeBound(surface.SpatialRoot));
            upper = CollisionSweep.Finite(minimum) ? (double?)minimum : null;
            lower = CollisionSweep.Finite(bound) ? (double?)Math.Min(bound, minimum) : null;
        }

        private static SpatialNode BuildSpatial(TargetMeshGeometry mesh, int[] triangles, int start, int count, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            double[] low = { double.PositiveInfinity, double.PositiveInfinity, double.PositiveInfinity };
            double[] high = { double.NegativeInfinity, double.NegativeInfinity, double.NegativeInfinity };
            for (int n = start; n < start + count; n++)
            {
                if ((n & 255) == 0) token.ThrowIfCancellationRequested();
                for (int k = 0; k < 3; k++)
                {
                    var p = mesh.Vertices[mesh.TriangleIndices[triangles[n]*3+k]];
                    low[0]=Math.Min(low[0],p.X); low[1]=Math.Min(low[1],p.Y); low[2]=Math.Min(low[2],p.Z);
                    high[0]=Math.Max(high[0],p.X); high[1]=Math.Max(high[1],p.Y); high[2]=Math.Max(high[2],p.Z);
                }
            }
            var node = new SpatialNode { Start=start, Count=count, Center=new BeamPoint3D((low[0]+high[0])/2,(low[1]+high[1])/2,(low[2]+high[2])/2) };
            node.Radius=Math.Sqrt(Distance2(node.Center,new BeamPoint3D(high[0],high[1],high[2]))) + 1e-8;
            if (count <= 24) return node;
            int axis = high[0]-low[0] >= high[1]-low[1] ? 0 : 1;
            if (high[2]-low[2] > high[axis]-low[axis]) axis=2;
            Func<int,double> center = tri => {
                var a=mesh.Vertices[mesh.TriangleIndices[tri*3]]; var b=mesh.Vertices[mesh.TriangleIndices[tri*3+1]]; var c=mesh.Vertices[mesh.TriangleIndices[tri*3+2]];
                return axis==0 ? a.X+b.X+c.X : axis==1 ? a.Y+b.Y+c.Y : a.Z+b.Z+c.Z; };
            Array.Sort(triangles,start,count,Comparer<int>.Create((a,b)=>center(a).CompareTo(center(b))));
            int half=count/2;
            node.Left=BuildSpatial(mesh,triangles,start,half,token); node.Right=BuildSpatial(mesh,triangles,start+half,count-half,token);
            return node;
        }

        internal static void Sample(PreparedSurface surface, Func<BeamPoint3D, string, double> distance, CancellationToken token,
            out double? lower, out double? upper)
        {
            var mesh = surface.Surface.Mesh; var values = new double[mesh.Vertices.Count];
            double minimum = double.PositiveInfinity, bound = double.PositiveInfinity;
            for (int i = 0; i < values.Length; i++)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                values[i] = distance(mesh.Vertices[i], surface.Surface.Role);
            }
            for (int i = 0; i < mesh.TriangleIndices.Length; i += 3)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                int ia = mesh.TriangleIndices[i], ib = mesh.TriangleIndices[i + 1], ic = mesh.TriangleIndices[i + 2];
                var a = mesh.Vertices[ia]; var b = mesh.Vertices[ib]; var c = mesh.Vertices[ic];
                var center = new BeamPoint3D((a.X + b.X + c.X) / 3, (a.Y + b.Y + c.Y) / 3, (a.Z + b.Z + c.Z) / 3);
                double value = distance(center, surface.Surface.Role);
                minimum = Math.Min(minimum, Math.Min(value, Math.Min(values[ia], Math.Min(values[ib], values[ic]))));
                // Every point of a triangle lies within its farthest-vertex radius from each sample.
                // Signed Euclidean distance is 1-Lipschitz: f(sample)-radius <= min_triangle f.
                double radius = Math.Sqrt(Math.Max(Distance2(center, a), Math.Max(Distance2(center, b), Distance2(center, c))));
                double ab = Math.Sqrt(Distance2(a, b)), ac = Math.Sqrt(Distance2(a, c)), bc = Math.Sqrt(Distance2(b, c));
                double triangleBound = Math.Max(value - radius, Math.Max(values[ia] - Math.Max(ab, ac),
                    Math.Max(values[ib] - Math.Max(ab, bc), values[ic] - Math.Max(ac, bc))));
                bound = Math.Min(bound, triangleBound);
            }
            upper = CollisionSweep.Finite(minimum) ? (double?)minimum : null;
            lower = CollisionSweep.Finite(bound) ? (double?)Math.Min(bound, minimum) : null;
        }

        internal static Func<BeamPoint3D, string, double> ProfilePointDistance(CollisionProfile profile, CollisionPose pose)
        {
            if (profile.Kind != "CArmSphere") return null; // Finite bore clipping requires the exact profile path.
            double x = pose.Source.X - pose.Isocenter.X, y = pose.Source.Y - pose.Isocenter.Y, z = pose.Source.Z - pose.Isocenter.Z;
            double scale = profile.HeadCenterFromIsoMm / Math.Sqrt(x * x + y * y + z * z);
            double cx = pose.Isocenter.X + scale * x, cy = pose.Isocenter.Y + scale * y, cz = pose.Isocenter.Z + scale * z;
            return (point, role) => { double dx = point.X - cx, dy = point.Y - cy, dz = point.Z - cz;
                return Math.Sqrt(dx * dx + dy * dy + dz * dz) - profile.HeadRadiusMm; };
        }

        internal sealed class SourceSolid
        {
            private SourceCollisionModel model;
            private readonly List<Segment> boundary = new List<Segment>();
            private readonly List<Arc> arcs = new List<Arc>();
            private sealed class Segment { public double X0, Y0, X1, Y1; }
            private sealed class Arc { public double Low, High; }

            internal static SourceSolid Create(SourceCollisionModel model)
            {
                if (model == null) return null;
                var result = new SourceSolid { model = model };
                if (model.Kind == "TrueBeamHeadSource")
                {
                    if (!Range(model.ReachRadiusMm, 1, 10000) || !Range(model.WarningMarginMm, 0, 1000) ||
                        !Range(model.SourceAngularMarginDegrees, 0, 90) || model.HeadObstacles == null || model.HeadObstacles.Count == 0 || model.HeadObstacles.Count > 16 ||
                        model.BoreLimitRadiusMm != 0 || model.BoreWarningRadiusMm != 0 || model.LeadingLimitFromIsoMm != 0 || model.LeadingWarningFromIsoMm != 0 ||
                        model.HeadObstacles.Any(h => h == null || !Range(h.FrontFaceFromIsoMm, 1, model.ReachRadiusMm) ||
                            !Range(h.RadiusMm, 1, model.ReachRadiusMm) || model.WarningMarginMm >= h.FrontFaceFromIsoMm)) return null;
                    result.BuildHeadBoundary();
                }
                else if (model.Kind == "HalcyonBoreSource")
                {
                    if (!Range(model.BoreLimitRadiusMm, 1, 10000) || !Range(model.BoreWarningRadiusMm, 1, model.BoreLimitRadiusMm) ||
                        model.BoreWarningRadiusMm >= model.BoreLimitRadiusMm || !Range(model.LeadingLimitFromIsoMm, 1, 10000) ||
                        model.ReachRadiusMm != 0 || model.WarningMarginMm != 0 || model.SourceAngularMarginDegrees != 0 ||
                        (model.HeadObstacles != null && model.HeadObstacles.Count != 0) ||
                        !Range(model.LeadingWarningFromIsoMm, 1, model.LeadingLimitFromIsoMm) || model.LeadingWarningFromIsoMm >= model.LeadingLimitFromIsoMm) return null;
                }
                else return null;
                return result;
            }

            private void BuildHeadBoundary()
            {
                // In the axial/radial half-plane, coaxial cylinders form a monotone staircase.
                // Only exposed front/side segments and reach-sphere arcs belong to the UNION boundary;
                // min(individual signed distances) would give a false penetration depth in overlaps.
                var steps = model.HeadObstacles.GroupBy(h => h.FrontFaceFromIsoMm).OrderBy(g => g.Key)
                    .Select(g => new SourceHeadObstacle { FrontFaceFromIsoMm = g.Key, RadiusMm = g.Max(h => h.RadiusMm) }).ToList();
                double radius = 0, reach = model.ReachRadiusMm;
                for (int i = 0; i < steps.Count; i++)
                {
                    double front = steps[i].FrontFaceFromIsoMm, end = i + 1 < steps.Count ? steps[i + 1].FrontFaceFromIsoMm : reach;
                    double sphereRadius = Math.Sqrt(Math.Max(0, reach * reach - front * front));
                    double old = Math.Min(radius, sphereRadius); radius = Math.Max(radius, steps[i].RadiusMm);
                    double updated = Math.Min(radius, sphereRadius);
                    if (updated > old || i == 0) boundary.Add(new Segment { X0 = front, Y0 = old, X1 = front, Y1 = updated });
                    double transition = Math.Sqrt(Math.Max(0, reach * reach - radius * radius));
                    if (front < Math.Min(end, transition)) boundary.Add(new Segment { X0 = front, Y0 = radius, X1 = Math.Min(end, transition), Y1 = radius });
                    double sphereStart = Math.Max(front, transition);
                    if (sphereStart <= end) arcs.Add(new Arc { Low = Math.Acos(Math.Min(1, end / reach)), High = Math.Acos(Math.Min(1, sphereStart / reach)) });
                }
            }

            internal Func<BeamPoint3D, string, double> AtPose(CollisionPose pose)
            {
                double ux = pose.Source.X - pose.Isocenter.X, uy = pose.Source.Y - pose.Isocenter.Y, uz = pose.Source.Z - pose.Isocenter.Z;
                double length = Math.Sqrt(ux * ux + uy * uy + uz * uz); ux /= length; uy /= length; uz /= length;
                return (point, role) =>
                {
                    double x = point.X - pose.Isocenter.X, y = point.Y - pose.Isocenter.Y, z = point.Z - pose.Isocenter.Z;
                    if (model.Kind == "HalcyonBoreSource")
                    {
                        double radialGap = model.BoreLimitRadiusMm - Math.Sqrt(x * x + y * y);
                        if (role != "External") return radialGap;
                        double leadingGap = model.LeadingLimitFromIsoMm - z;
                        // Distance to the complement (safe cylinder intersected with the leading half-space).
                        if (radialGap < 0 && leadingGap < 0) return -Math.Sqrt(radialGap * radialGap + leadingGap * leadingGap);
                        return Math.Min(radialGap, leadingGap);
                    }
                    double along = x * ux + y * uy + z * uz;
                    double rx = x - along * ux, ry = y - along * uy, rz = z - along * uz;
                    return HeadDistance(along, Math.Sqrt(rx * rx + ry * ry + rz * rz));
                };
            }

            private double HeadDistance(double axial, double radial)
            {
                double squared = double.PositiveInfinity;
                foreach (var segment in boundary)
                {
                    double x = segment.X1 - segment.X0, y = segment.Y1 - segment.Y0, norm = x * x + y * y;
                    double t = norm == 0 ? 0 : Math.Max(0, Math.Min(1, ((axial - segment.X0) * x + (radial - segment.Y0) * y) / norm));
                    double dx = axial - segment.X0 - t * x, dy = radial - segment.Y0 - t * y;
                    squared = Math.Min(squared, dx * dx + dy * dy);
                }
                foreach (var arc in arcs)
                {
                    double theta = Math.Max(arc.Low, Math.Min(arc.High, Math.Atan2(radial, axial)));
                    double dx = axial - model.ReachRadiusMm * Math.Cos(theta), dy = radial - model.ReachRadiusMm * Math.Sin(theta);
                    squared = Math.Min(squared, dx * dx + dy * dy);
                }
                bool inside = axial * axial + radial * radial <= model.ReachRadiusMm * model.ReachRadiusMm &&
                    model.HeadObstacles.Any(h => axial >= h.FrontFaceFromIsoMm && radial <= h.RadiusMm);
                return (inside ? -1 : 1) * Math.Sqrt(squared);
            }
            private static bool Range(double value, double min, double max) { return CollisionSweep.Finite(value) && value >= min && value <= max; }
        }

        private static double Distance2(BeamPoint3D a, BeamPoint3D b)
        { double x = a.X - b.X, y = a.Y - b.Y, z = a.Z - b.Z; return x * x + y * y + z * z; }
    }
}
