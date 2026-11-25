using System;
using System.Collections.Generic;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    public sealed class SourceCollisionScreeningRequest
    {
        public string MachineId { get; set; }
        public string CoordinateSystem { get; set; }
        public string PatientPosition { get; set; }
        public bool NativeSourceCoordinates { get; set; }
        public List<SourceCollisionSurface> Surfaces { get; set; }
        public List<CollisionPose> Poses { get; set; }
    }

    public sealed class SourceCollisionSurface
    {
        public string Label { get; set; }
        public string Role { get; set; }
        public List<BeamPoint3D> Points { get; set; }
    }

    public sealed class SourceCollisionScreeningResult
    {
        public SourceCollisionScreeningResult() { Findings = new List<SourceCollisionScreeningFinding>(); }
        public bool ClinicalClearanceSupported { get { return false; } }
        public string Status { get; set; }
        public string Summary { get; set; }
        public string ModelId { get; set; }
        public string SourceSha256 { get; set; }
        public string EvidenceLevel { get; set; }
        public List<SourceCollisionScreeningFinding> Findings { get; set; }
    }

    public sealed class SourceCollisionScreeningFinding
    {
        public string BeamId { get; set; }
        public int ControlPointIndex { get; set; }
        public string SurfaceLabel { get; set; }
        public int ScreenedPointCount { get; set; }
        public int ExcludedPointCount { get; set; }
        public int ModelHitPointCount { get; set; }
        public int MarginHitPointCount { get; set; }
    }

    public static class SourceCollisionScreening
    {
        public const int MaximumPoints = 200000;
        public const long MaximumPointPosePairs = 100000000;
        public const int MaximumFindings = 10000;
        private const string Scope = "Sampled surface-point screening at captured poses only; not clinical clearance. " +
            "No triangle intersection, minimum surface distance, continuous trajectory, missing anatomy or unsampled accessory coverage is established.";

        public static SourceCollisionScreeningResult Evaluate(SourceCollisionScreeningRequest request,
            SourceCollisionModelCatalog catalog, CancellationToken cancellationToken)
        {
            try
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (catalog == null) return Unavailable("Source model catalog is missing.");
                var model = catalog.Select(request == null ? null : request.MachineId);
                if (model == null) return Unavailable("No source model is bound to the exact native machine ID.");
                if (!Valid(request, model, cancellationToken)) return Unavailable("Geometry, native coordinates, pose or supported orientation is missing, invalid or exceeds screening bounds.");
                var result = new SourceCollisionScreeningResult { Status = "screening-no-hit", Summary = Scope,
                    ModelId = model.ModelId, SourceSha256 = model.SourceSha256, EvidenceLevel = model.EvidenceLevel };
                if (!request.Surfaces.Exists(s => s.Role == "external")) result.Summary += " Only supplied surface points screened: external missing.";
                if (!request.Surfaces.Exists(s => s.Role == "support")) result.Summary += " Only supplied surface points screened: support missing.";
                bool anyHit = false, uncovered = false;
                foreach (var pose in request.Poses)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    double ux = pose.Source.X - pose.Isocenter.X, uy = pose.Source.Y - pose.Isocenter.Y, uz = pose.Source.Z - pose.Isocenter.Z;
                    double length = Math.Sqrt(ux * ux + uy * uy + uz * uz);
                    ux /= length; uy /= length; uz /= length;
                    foreach (var surface in request.Surfaces)
                    {
                        var finding = new SourceCollisionScreeningFinding { BeamId = pose.BeamId,
                            ControlPointIndex = pose.ControlPointIndex, SurfaceLabel = surface.Label };
                        for (int i = 0; i < surface.Points.Count; i++)
                        {
                            if ((i & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                            var point = surface.Points[i];
                            double x = point.X - pose.Isocenter.X, y = point.Y - pose.Isocenter.Y, z = point.Z - pose.Isocenter.Z;
                            bool hit, marginHit;
                            if (model.Kind == "TrueBeamHeadSource")
                            {
                                double distance2 = x * x + y * y + z * z;
                                if (!AtMost(distance2, model.ReachRadiusMm * model.ReachRadiusMm))
                                { finding.ExcludedPointCount++; continue; }
                                double along = x * ux + y * uy + z * uz;
                                // Use perpendicular-vector length, avoiding cancellation in distance2 - along*along.
                                double rx = x - along * ux, ry = y - along * uy, rz = z - along * uz;
                                double radial2 = rx * rx + ry * ry + rz * rz;
                                hit = HeadHit(model, along, radial2, 0);
                                marginHit = HeadHit(model, along, radial2, model.WarningMarginMm);
                            }
                            else
                            {
                                double radial2 = x * x + y * y;
                                hit = AtMost(model.BoreLimitRadiusMm * model.BoreLimitRadiusMm, radial2) ||
                                    (surface.Role == "external" && AtMost(model.LeadingLimitFromIsoMm, z));
                                marginHit = AtMost(model.BoreWarningRadiusMm * model.BoreWarningRadiusMm, radial2) ||
                                    (surface.Role == "external" && AtMost(model.LeadingWarningFromIsoMm, z));
                            }
                            finding.ScreenedPointCount++;
                            if (hit) finding.ModelHitPointCount++;
                            if (marginHit) finding.MarginHitPointCount++;
                        }
                        anyHit |= finding.MarginHitPointCount > 0 || finding.ModelHitPointCount > 0;
                        uncovered |= finding.ScreenedPointCount == 0;
                        result.Findings.Add(finding);
                    }
                }
                cancellationToken.ThrowIfCancellationRequested();
                result.Status = uncovered ? "unavailable" : anyHit ? "screening-hit" : "screening-no-hit";
                if (uncovered) result.Summary = "A surface/pose has no points within the source model reach; other findings are retained. " + result.Summary;
                return result;
            }
            catch (OperationCanceledException)
            { return new SourceCollisionScreeningResult { Status = "cancelled", Summary = "Screening cancelled; partial findings discarded. " + Scope }; }
            catch (ArgumentException)
            { return Unavailable("Source model catalog is invalid; no commissioning or coordinate assumption is inferred."); }
        }

        private static bool HeadHit(SourceCollisionModel model, double along, double radial2, double margin)
        {
            foreach (var obstacle in model.HeadObstacles)
            {
                double radius = obstacle.RadiusMm + margin;
                if (AtMost(obstacle.FrontFaceFromIsoMm - margin, along) && AtMost(radial2, radius * radius)) return true;
            }
            return false;
        }

        private static bool Valid(SourceCollisionScreeningRequest request, SourceCollisionModel model, CancellationToken cancellationToken)
        {
            if (request == null || !request.NativeSourceCoordinates || request.CoordinateSystem != "DICOM-LPS-mm" ||
                (request.PatientPosition != "HFS" && request.PatientPosition != "FFS" && request.PatientPosition != "HFP" && request.PatientPosition != "FFP") ||
                (model.Kind == "HalcyonBoreSource" && request.PatientPosition != "HFS") ||
                request.Poses == null || request.Poses.Count == 0 || request.Poses.Count > 10000 ||
                request.Surfaces == null || request.Surfaces.Count == 0 || request.Surfaces.Count > 64 ||
                (long)request.Surfaces.Count * request.Poses.Count > MaximumFindings) return false;
            long points = 0;
            var surfaceLabels = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var surface in request.Surfaces)
            {
                if (surface == null || !SourceCollisionModelCatalog.Text(surface.Label, 256) ||
                    !surfaceLabels.Add(surface.Label) ||
                    (surface.Role != "external" && surface.Role != "support") || surface.Points == null || surface.Points.Count == 0) return false;
                points += surface.Points.Count;
                if (points > MaximumPoints || points * request.Poses.Count > MaximumPointPosePairs) return false;
                for (int i = 0; i < surface.Points.Count; i++)
                {
                    if ((i & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
                    if (!Finite(surface.Points[i])) return false;
                }
            }
            var poseIndices = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            foreach (var pose in request.Poses)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (pose == null || !SourceCollisionModelCatalog.Text(pose.BeamId, 256) || pose.ControlPointIndex < 0 || pose.ControlPointIndex > 1000000 ||
                    !Finite(pose.Isocenter) || !Finite(pose.Source) || double.IsNaN(pose.GantryDegrees) || pose.GantryDegrees < 0 || pose.GantryDegrees > 360) return false;
                HashSet<int> indices;
                if (!poseIndices.TryGetValue(pose.BeamId, out indices))
                { indices = new HashSet<int>(); poseIndices.Add(pose.BeamId, indices); }
                if (!indices.Add(pose.ControlPointIndex)) return false;
                double x = pose.Source.X - pose.Isocenter.X, y = pose.Source.Y - pose.Isocenter.Y, z = pose.Source.Z - pose.Isocenter.Z;
                if (x * x + y * y + z * z < 1e-12) return false;
            }
            return true;
        }

        private static bool Finite(BeamPoint3D point)
        { return point != null && Bounded(point.X) && Bounded(point.Y) && Bounded(point.Z); }
        private static bool Bounded(double value)
        { return !double.IsNaN(value) && value >= -1000000 && value <= 1000000; }
        // Relative floating-point guard only, not a physical safety margin.
        private static bool AtMost(double value, double boundary)
        { return value <= boundary || value - boundary <= 1e-12 * Math.Max(1.0, Math.Max(Math.Abs(value), Math.Abs(boundary))); }
        private static SourceCollisionScreeningResult Unavailable(string reason)
        { return new SourceCollisionScreeningResult { Status = "unavailable", Summary = reason + " " + Scope }; }
    }
}
