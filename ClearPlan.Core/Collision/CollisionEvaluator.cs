using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    public static class CollisionEvaluator
    {
        private const double Epsilon = 1e-8;
        private const string Limitations = " Research preview: sampled poses only; no arc or interfield interpolation. " +
            "CT truncation, missing accessories and actual setup remain unverified. No insertion path is evaluated. " +
            "Mesh self-intersections and source-surface accuracy remain unverified. " +
            "Envelope and mesh assumptions require local validation; this is not clinical clearance.";

        public static CollisionResult Evaluate(CollisionScene scene, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                ValidateScene(scene, cancellationToken);
                var result = new CollisionResult();
                foreach (var pose in scene.Poses) foreach (var surface in scene.Surfaces)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    double? clearance = scene.Profile.Kind == "CArmSphere"
                        ? SphereClearance(surface.Mesh, pose, scene.Profile, cancellationToken)
                        : BoreClearance(surface.Mesh, pose.Isocenter, scene.Profile, cancellationToken);
                    if (clearance.HasValue && !Finite(clearance.Value)) clearance = null;
                    string status = !clearance.HasValue ? "unavailable" :
                        clearance.Value <= 0 ? "overlap" : clearance.Value <= scene.Profile.SafetyMarginMm ? "near" : "sampled-clear";
                    result.Findings.Add(new CollisionFinding { BeamId = pose.BeamId, ControlPointIndex = pose.ControlPointIndex,
                        GantryDegrees = pose.GantryDegrees, SurfaceLabel = surface.Label, ClearanceMm = clearance, Status = status,
                        Message = !clearance.HasValue ? "No axial surface coverage or ambiguous mesh containment; no clearance available."
                            : "Sampled envelope-to-surface clearance only; no continuous-motion or clinical clearance." });
                }
                result.Status = result.Findings.Any(f => f.Status == "unavailable") ? "unavailable" :
                    result.Findings.Any(f => f.Status == "overlap") ? "overlap" :
                    result.Findings.Any(f => f.Status == "near") ? "near" : "sampled-clear";
                result.Summary = (scene.Synthetic ? "SYNTHETIC demonstration. " : "Commissioned-envelope research evaluation. ") +
                    "Evaluated " + scene.Poses.Count + " sampled poses and " + scene.Surfaces.Count + " surfaces. " +
                    (result.Status == "unavailable" ? "Coverage incomplete; inspect individual findings. " : "Envelope result: " + result.Status + ". ") +
                    (scene.Profile.Kind == "RingBore" ? "Finite patient-z bore slab assumes HFS and zero couch. " : "Conservative full head sphere. ") + Limitations;
                return result;
            }
            catch (ArgumentException error) { return new CollisionResult { Status = "unavailable", Summary = error.Message + Limitations }; }
        }

        private static void ValidateScene(CollisionScene scene, CancellationToken token)
        {
            if (scene == null || !scene.SupportedCoordinates) Invalid("Supported and verified patient coordinates are required.");
            CollisionProfileCatalog.Validate(scene.Profile, scene.Synthetic);
            if (scene.Surfaces == null || scene.Surfaces.Count < 2 || scene.Surfaces.Count > 32 ||
                scene.Surfaces.Any(s => s == null || string.IsNullOrWhiteSpace(s.Label) || s.Label.Length > 256 ||
                    (s.Role != "External" && s.Role != "Support")) ||
                !scene.Surfaces.Any(s => s.Role == "External") || !scene.Surfaces.Any(s => s.Role == "Support"))
                Invalid("Nonempty External and Support surfaces are both required.");
            if (scene.Poses == null || scene.Poses.Count == 0 || scene.Poses.Count > 10000) Invalid("Sampled poses are missing or exceed the preview budget.");
            long triangles = 0;
            foreach (var surface in scene.Surfaces)
            {
                if (surface.Mesh == null || surface.Mesh.TriangleIndices == null) Invalid("A closed triangle mesh is missing.");
                triangles += surface.Mesh.TriangleIndices.Length / 3;
            }
            // Distance plus three containment rays, with bounded retries at edges/vertices.
            if (triangles * scene.Poses.Count > 10000000L) Invalid("Geometry and pose workload exceeds the bounded preview budget.");
            foreach (var surface in scene.Surfaces) ValidateMesh(surface.Mesh, token);
            var identities = new HashSet<string>(StringComparer.Ordinal);
            foreach (var pose in scene.Poses)
            {
                token.ThrowIfCancellationRequested();
                if (pose == null || string.IsNullOrWhiteSpace(pose.BeamId) || pose.BeamId.Length > 256 || pose.ControlPointIndex < 0 ||
                    !Finite(pose.GantryDegrees) || Math.Abs(pose.GantryDegrees) > 3600 || !PointValid(pose.Isocenter) || !PointValid(pose.Source) ||
                    LengthSquared(Sub(pose.Source, pose.Isocenter)) < Epsilon * Epsilon ||
                    !identities.Add(pose.BeamId + "\n" + pose.ControlPointIndex)) Invalid("A sampled pose is invalid, duplicated or lacks source/isocenter geometry.");
            }
        }

        private static void ValidateMesh(TargetMeshGeometry mesh, CancellationToken token)
        {
            if (mesh == null || mesh.Vertices == null || mesh.TriangleIndices == null || mesh.Vertices.Count < 4 ||
                mesh.Vertices.Count > 250000 || mesh.TriangleIndices.Length < 12 || mesh.TriangleIndices.Length > 1500000 ||
                mesh.TriangleIndices.Length % 3 != 0) Invalid("A closed triangle mesh is missing, empty or exceeds the bounded preview budget.");
            for (int i = 0; i < mesh.Vertices.Count; i++)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                if (!PointValid(mesh.Vertices[i])) Invalid("Mesh coordinates must be finite and within the bounded geometry range.");
            }
            var edges = new Dictionary<long, EdgeUse>();
            for (int i = 0; i < mesh.TriangleIndices.Length; i += 3)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                int a = mesh.TriangleIndices[i], b = mesh.TriangleIndices[i+1], c = mesh.TriangleIndices[i+2];
                if (a < 0 || b < 0 || c < 0 || a >= mesh.Vertices.Count || b >= mesh.Vertices.Count || c >= mesh.Vertices.Count || a == b || b == c || a == c)
                    Invalid("Mesh triangle indices are invalid.");
                var ab = Sub(mesh.Vertices[b],mesh.Vertices[a]); var ac = Sub(mesh.Vertices[c],mesh.Vertices[a]);
                double longestSquared = Math.Max(LengthSquared(ab),Math.Max(LengthSquared(ac),LengthSquared(Sub(mesh.Vertices[c],mesh.Vertices[b]))));
                double areaSquared = LengthSquared(Cross(ab,ac));
                if (areaSquared <= Epsilon * Epsilon || areaSquared <= longestSquared*longestSquared*1e-12)
                    Invalid("Mesh contains a degenerate or numerically ill-conditioned triangle.");
                AddEdge(edges, a, b); AddEdge(edges, b, c); AddEdge(edges, c, a);
            }
            if (edges.Values.Any(e => e.Count != 2 || e.Orientation != 0)) Invalid("Mesh must be watertight with consistent edge orientation; open or nonmanifold edges are unsupported.");
        }

        private sealed class EdgeUse { public int Count; public int Orientation; }
        private static void AddEdge(Dictionary<long, EdgeUse> edges, int a, int b)
        {
            long key = ((long)Math.Min(a,b) << 32) | (uint)Math.Max(a,b);
            EdgeUse edge;
            if (!edges.TryGetValue(key, out edge)) { edge = new EdgeUse(); edges.Add(key, edge); }
            edge.Count++; edge.Orientation += a < b ? 1 : -1;
            if (edge.Count > 2) Invalid("Mesh contains a nonmanifold or duplicate triangle edge.");
        }

        private static double? SphereClearance(TargetMeshGeometry mesh, CollisionPose pose, CollisionProfile profile, CancellationToken token)
        {
            var direction = Sub(pose.Source, pose.Isocenter);
            var center = Add(pose.Isocenter, Scale(direction, profile.HeadCenterFromIsoMm / Math.Sqrt(LengthSquared(direction))));
            double distanceSquared = double.PositiveInfinity;
            for (int i = 0; i < mesh.TriangleIndices.Length; i += 3)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                distanceSquared = Math.Min(distanceSquared, PointTriangleDistanceSquared(center,
                    mesh.Vertices[mesh.TriangleIndices[i]], mesh.Vertices[mesh.TriangleIndices[i+1]], mesh.Vertices[mesh.TriangleIndices[i+2]]));
            }
            double distance = Math.Sqrt(distanceSquared);
            if (distance <= Epsilon) return -profile.HeadRadiusMm;
            bool? inside = Contains(mesh, center, token);
            return inside.HasValue ? (double?)((inside.Value ? -distance : distance) - profile.HeadRadiusMm) : null;
        }

        // Voronoi regions of the triangle: face interiors, edges and vertices all contribute.
        private static double PointTriangleDistanceSquared(BeamPoint3D p, BeamPoint3D a, BeamPoint3D b, BeamPoint3D c)
        {
            var ab = Sub(b,a); var ac = Sub(c,a); var ap = Sub(p,a);
            double d1 = Dot(ab,ap), d2 = Dot(ac,ap);
            if (d1 <= 0 && d2 <= 0) return LengthSquared(ap);
            var bp = Sub(p,b); double d3 = Dot(ab,bp), d4 = Dot(ac,bp);
            if (d3 >= 0 && d4 <= d3) return LengthSquared(bp);
            double vc = d1*d4-d3*d2;
            if (vc <= 0 && d1 >= 0 && d3 <= 0) return LengthSquared(Sub(p,Add(a,Scale(ab,d1/(d1-d3)))));
            var cp = Sub(p,c); double d5 = Dot(ab,cp), d6 = Dot(ac,cp);
            if (d6 >= 0 && d5 <= d6) return LengthSquared(cp);
            double vb = d5*d2-d1*d6;
            if (vb <= 0 && d2 >= 0 && d6 <= 0) return LengthSquared(Sub(p,Add(a,Scale(ac,d2/(d2-d6)))));
            double va = d3*d6-d5*d4;
            if (va <= 0 && d4-d3 >= 0 && d5-d6 >= 0)
                return LengthSquared(Sub(p,Add(b,Scale(Sub(c,b),(d4-d3)/((d4-d3)+(d5-d6))))));
            double denominator = 1/(va+vb+vc);
            return LengthSquared(Sub(p,Add(a,Add(Scale(ab,vb*denominator),Scale(ac,vc*denominator)))));
        }

        private static bool? Contains(TargetMeshGeometry mesh, BeamPoint3D center, CancellationToken token)
        {
            // Edge/vertex and coplanar rays are discarded, not counted twice. Three independent
            // unambiguous rays must agree; unresolved or inconsistent parity fails closed.
            var directions = new[] { new BeamPoint3D(1,0.371,0.529), new BeamPoint3D(0.217,1,0.613),
                new BeamPoint3D(0.419,0.733,1), new BeamPoint3D(1,-0.327,0.811), new BeamPoint3D(-0.547,1,0.293),
                new BeamPoint3D(0.683,-0.443,1), new BeamPoint3D(1,0.887,-0.199) };
            bool? answer = null; int agreeing = 0;
            foreach (var direction in directions)
            {
                var hits = new List<double>(); bool ambiguous = false;
                for (int i = 0; i < mesh.TriangleIndices.Length; i += 3)
                {
                    if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                    RayHit(center, direction, mesh.Vertices[mesh.TriangleIndices[i]], mesh.Vertices[mesh.TriangleIndices[i+1]],
                        mesh.Vertices[mesh.TriangleIndices[i+2]], hits, ref ambiguous);
                    if (ambiguous) break;
                }
                if (ambiguous) continue;
                hits.Sort(); int unique = 0; double previous = double.NegativeInfinity;
                foreach (double hit in hits)
                    if (hit-previous > Epsilon * Math.Max(1,Math.Abs(hit))) { unique++; previous = hit; }
                bool inside = (unique & 1) != 0;
                if (answer.HasValue && answer.Value != inside) return null;
                answer = inside;
                if (++agreeing == 3) return answer;
            }
            return null;
        }

        private static void RayHit(BeamPoint3D p, BeamPoint3D direction, BeamPoint3D a, BeamPoint3D b, BeamPoint3D c,
            List<double> hits, ref bool ambiguous)
        {
            var ab = Sub(b,a); var ac = Sub(c,a); var normal = Cross(ab,ac); var pa = Sub(p,a);
            var cross = Cross(direction,ac); double determinant = Dot(ab,cross);
            if (Math.Abs(determinant) <= 1e-12 * Math.Sqrt(LengthSquared(normal)*LengthSquared(direction)))
            {
                if (Math.Abs(Dot(pa,normal)) <= Epsilon*Math.Sqrt(LengthSquared(normal))) ambiguous = true;
                return;
            }
            double inverse = 1/determinant, u = Dot(pa,cross)*inverse;
            var q = Cross(pa,ab); double v = Dot(direction,q)*inverse, t = Dot(ac,q)*inverse;
            if (u < -Epsilon || v < -Epsilon || u+v > 1+Epsilon || t <= Epsilon) return;
            if (u <= Epsilon || v <= Epsilon || 1-u-v <= Epsilon) { ambiguous = true; return; }
            hits.Add(t);
        }

        private static double? BoreClearance(TargetMeshGeometry mesh, BeamPoint3D iso, CollisionProfile profile, CancellationToken token)
        {
            double maximumSquared = -1;
            for (int i = 0; i < mesh.TriangleIndices.Length; i += 3)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                var polygon = new List<BeamPoint3D> { Sub(mesh.Vertices[mesh.TriangleIndices[i]],iso),
                    Sub(mesh.Vertices[mesh.TriangleIndices[i+1]],iso), Sub(mesh.Vertices[mesh.TriangleIndices[i+2]],iso) };
                polygon = ClipZ(ClipZ(polygon, -profile.BoreHalfLengthMm, true), profile.BoreHalfLengthMm, false);
                foreach (var vertex in polygon) maximumSquared = Math.Max(maximumSquared, vertex.X*vertex.X + vertex.Y*vertex.Y);
            }
            return maximumSquared < 0 ? null : (double?)(profile.BoreRadiusMm - Math.Sqrt(maximumSquared));
        }

        private static List<BeamPoint3D> ClipZ(List<BeamPoint3D> polygon, double boundary, bool lower)
        {
            var clipped = new List<BeamPoint3D>();
            if (polygon.Count == 0) return clipped;
            var previous = polygon[polygon.Count-1]; bool previousInside = lower ? previous.Z >= boundary : previous.Z <= boundary;
            foreach (var current in polygon)
            {
                bool inside = lower ? current.Z >= boundary : current.Z <= boundary;
                if (inside != previousInside)
                    clipped.Add(Add(previous,Scale(Sub(current,previous),(boundary-previous.Z)/(current.Z-previous.Z))));
                if (inside) clipped.Add(current);
                previous = current; previousInside = inside;
            }
            return clipped;
        }

        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool PointValid(BeamPoint3D p) { return p != null && Finite(p.X) && Finite(p.Y) && Finite(p.Z) && Math.Abs(p.X) <= 1000000 && Math.Abs(p.Y) <= 1000000 && Math.Abs(p.Z) <= 1000000; }
        private static BeamPoint3D Sub(BeamPoint3D a, BeamPoint3D b) { return new BeamPoint3D(a.X-b.X,a.Y-b.Y,a.Z-b.Z); }
        private static BeamPoint3D Add(BeamPoint3D a, BeamPoint3D b) { return new BeamPoint3D(a.X+b.X,a.Y+b.Y,a.Z+b.Z); }
        private static BeamPoint3D Scale(BeamPoint3D a, double value) { return new BeamPoint3D(a.X*value,a.Y*value,a.Z*value); }
        private static double Dot(BeamPoint3D a, BeamPoint3D b) { return a.X*b.X+a.Y*b.Y+a.Z*b.Z; }
        private static double LengthSquared(BeamPoint3D a) { return Dot(a,a); }
        private static BeamPoint3D Cross(BeamPoint3D a, BeamPoint3D b) { return new BeamPoint3D(a.Y*b.Z-a.Z*b.Y,a.Z*b.X-a.X*b.Z,a.X*b.Y-a.Y*b.X); }
        private static void Invalid(string message) { throw new ArgumentException(message); }
    }
}
