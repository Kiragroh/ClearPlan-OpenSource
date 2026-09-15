using System;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    /// <summary>CT acquisition extent, independent of whether the TPS capped the External mesh.</summary>
    public static class CollisionCtCoverage
    {
        public static void Apply(CollisionScene scene, BeamPoint3D origin, BeamPoint3D[] axes, double[] spacing, int[] sizes)
        {
            if (scene == null) throw new ArgumentNullException(nameof(scene));
            scene.CtCoverageKnown = false;
            scene.ExternalTruncatedAtCtBoundary = false;
            scene.CtCoverageReason = "CT extent unavailable or invalid; missing anatomy cannot be excluded.";
            if (!Valid(origin) || axes == null || axes.Length != 3 || spacing == null || spacing.Length != 3 || sizes == null || sizes.Length != 3) return;
            for (int axis = 0; axis < 3; axis++)
            {
                if (!Valid(axes[axis]) || Math.Abs(Dot(axes[axis], axes[axis]) - 1) > 1e-5 ||
                    !Finite(spacing[axis]) || spacing[axis] <= 0 || spacing[axis] > 1000 || sizes[axis] < 2 || sizes[axis] > 1000000) return;
                for (int other = 0; other < axis; other++) if (Math.Abs(Dot(axes[axis], axes[other])) > 1e-5) return;
            }
            var external = scene.Surfaces == null ? null : scene.Surfaces.Where(s => s != null && s.Role == "External").ToArray();
            if (external == null || external.Length != 1 || external[0].Mesh == null || external[0].Mesh.Vertices == null || external[0].Mesh.Vertices.Count == 0) return;
            bool boundary = false;
            foreach (var point in external[0].Mesh.Vertices)
            {
                if (!Valid(point)) return;
                var offset = new BeamPoint3D(point.X - origin.X, point.Y - origin.Y, point.Z - origin.Z);
                for (int axis = 0; axis < 3; axis++)
                {
                    double index = Dot(offset, axes[axis]) / spacing[axis];
                    // Conservatively include half a voxel around the first/last image center.
                    // Artificial caps on first/last CT slices therefore cannot imply full coverage.
                    if (index <= 0.5 || index >= sizes[axis] - 1.5) boundary = true;
                }
            }
            scene.CtCoverageKnown = true;
            scene.ExternalTruncatedAtCtBoundary = boundary;
            scene.CtCoverageReason = boundary
                ? "External reaches the CT boundary; anatomy beyond the acquired image is unknown."
                : "External does not reach the acquired CT boundary; unsampled accessories and actual setup remain outside scope.";
        }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
        private static bool Valid(BeamPoint3D p) { return p != null && Finite(p.X) && Finite(p.Y) && Finite(p.Z); }
        private static double Dot(BeamPoint3D a, BeamPoint3D b) { return a.X * b.X + a.Y * b.Y + a.Z * b.Z; }
    }
}
