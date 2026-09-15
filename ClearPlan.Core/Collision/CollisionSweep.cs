using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    public sealed class CollisionSweepResult
    {
        public CollisionSweepResult() { Frames = new List<CollisionSweepFrame>(); }
        public List<CollisionSweepFrame> Frames { get; set; }
        public string Status { get; set; }
        public string Summary { get; set; }
    }

    public sealed class CollisionSweepFrame
    {
        public int Index { get; set; }
        public CollisionPose Pose { get; set; }
        public bool Interpolated { get; set; }
        // Sweep ordinals are not native CP identities. Preserve the captured reference
        // before renumbering the derived rows; no interpolated aperture is invented.
        public int? CapturedControlPointIndex { get; set; }
        public double CapturedGantryDegrees { get; set; }
        public double CapturedCouchDegrees { get; set; }
        public string ModelBodyStatus { get; set; }
        public string ModelTableStatus { get; set; }
        public string Status { get; set; }
        public string Reason { get; set; }
        /// <summary>Profile result or sampled upper bound; RingBore supplies a radial gap only, not finite-end clearance. See Reason.</summary>
        public double? BodyClearanceMm { get; set; }
        public double? TableClearanceMm { get; set; }
        public double? BodyLowerBoundMm { get; set; }
        public double? TableLowerBoundMm { get; set; }
        public string BodyStatus { get; set; }
        public string TableStatus { get; set; }
        public CollisionRadialReview RadialReview { get; set; }
    }

    /// <summary>Distance to an open cylindrical bore for supplied surfaces only; not clinical clearance.</summary>
    public sealed class CollisionRadialReview
    {
        public double? BodyClearanceMm { get; set; }
        public double? TableClearanceMm { get; set; }
        public string BodyStatus { get; set; }
        public string TableStatus { get; set; }
        public string Status { get; set; }
        public string Scope { get; set; }
    }

    /// <summary>Detached, bounded research frames. Ten-degree states do not establish continuous-path or clinical clearance.</summary>
    public static class CollisionSweep
    {
        public const int MaximumInputPoses = 10000;
        public const int MaximumFrames = 2000;
        public const long MaximumTrianglePosePairs = 10000000;
        private const string Scope = "Research only; 10-degree states plus arc endpoints, not continuous-path or clinical clearance. Missing anatomy, accessories, setup and mesh self-intersections remain unverified.";
        private const double Epsilon = 1e-8;

        public static CollisionSweepResult Evaluate(CollisionScene scene, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                ValidateInputs(scene, token);
                var result = new CollisionSweepResult { Frames = SampleFrames(scene.Poses, token) };
                var surfaces = scene.Surfaces.Select(s => CollisionSweepDistance.Prepare(s, token)).ToList();
                long triangleCount = surfaces.Sum(s => s.Valid ? (long)s.Surface.Mesh.TriangleIndices.Length / 3 : 0);
                long vertexCount = surfaces.Sum(s => s.Valid ? (long)s.Surface.Mesh.Vertices.Count : 0);
                bool profile = false;
                string profileReason = null, modelReason = null;
                CollisionSweepDistance.SourceSolid solid = null;
                if (scene.Profile != null)
                {
                    try
                    {
                        CollisionProfileCatalog.Validate(scene.Profile, scene.Synthetic);
                        profile = scene.SupportedCoordinates;
                        if (!profile) profileReason = "Profile envelope is not selected because its coordinates are unconfirmed.";
                    }
                    catch (ArgumentException error) { profileReason = "Profile envelope is not selected: " + error.Message; }
                }
                if (!profile)
                {
                    solid = scene.Synthetic ? null : CollisionSweepDistance.SourceSolid.Create(scene.SourceModel);
                    if (solid == null) modelReason = "Source no-fly model is missing or has invalid bounded geometry.";
                    else if ((scene.SourceModel.Kind == "HalcyonBoreSource" && scene.PatientPosition != "HFS") ||
                        (scene.SourceModel.Kind == "TrueBeamHeadSource" && scene.PatientPosition != "HFS" && scene.PatientPosition != "HFP" &&
                         scene.PatientPosition != "FFS" && scene.PatientPosition != "FFP"))
                        modelReason = "Patient orientation is unsupported by the nominal source model.";
                }

                // A source Halcyon bore is stationary in patient coordinates at a fixed
                // isocenter. Its analytic solid does not depend on source/gantry direction.
                // Reuse only these identical distance evaluations; keep every angle and
                // its own coverage/coordinate/arc classification. Never simplify the mesh.
                bool stationaryBore = !profile && solid != null && modelReason == null && scene.SourceModel.Kind == "HalcyonBoreSource";
                bool spatialHead = !profile && solid != null && modelReason == null && scene.SourceModel.Kind == "TrueBeamHeadSource";
                int evaluations = stationaryBore ? result.Frames.Select(f => BoreKey(f.Pose)).Distinct().Count() : result.Frames.Count;
                bool distancePossible = (profile || (solid != null && modelReason == null)) && (scene.SupportedCoordinates || scene.NominalCoordinatesSupported);
                if ((distancePossible && !spatialHead && (triangleCount * evaluations > MaximumTrianglePosePairs || vertexCount * evaluations > MaximumTrianglePosePairs)) ||
                    triangleCount > MaximumTrianglePosePairs || vertexCount > MaximumTrianglePosePairs ||
                    (long)surfaces.Count * result.Frames.Count > 20000)
                    return Uncertain("Geometry and sampled-pose workload exceeds the bounded sweep budget.");
                var spatialBudget = new CollisionSweepDistance.DistanceBudget();
                var stationaryDistances = new Dictionary<Tuple<double, double, double, string>, Tuple<double?, double?>>();
                var radialReviews = new Dictionary<Tuple<double, double, double>, CollisionRadialReview>();

                // The existing commissioned-profile evaluator is called once for the complete batch:
                // its commissioning, topology and coordinate gates are never weakened or supplied by source screening.
                var exact = new Dictionary<string, List<CollisionFinding>>(StringComparer.Ordinal);
                if (profile && modelReason == null && scene.SupportedCoordinates && surfaces.All(s => s.Valid && s.Closed) &&
                    surfaces.Any(s => s.Surface.Role == "External") && surfaces.Any(s => s.Surface.Role == "Support"))
                {
                    var batch = CollisionEvaluator.Evaluate(new CollisionScene { Synthetic = scene.Synthetic,
                        SupportedCoordinates = scene.SupportedCoordinates, Profile = scene.Profile, Surfaces = scene.Surfaces,
                        Poses = result.Frames.Select(f => f.Pose).ToList() }, token);
                    foreach (var finding in batch.Findings)
                    {
                        string key = Key(finding.BeamId, finding.ControlPointIndex);
                        List<CollisionFinding> list;
                        if (!exact.TryGetValue(key, out list)) { list = new List<CollisionFinding>(); exact.Add(key, list); }
                        list.Add(finding);
                    }
                }

                foreach (var frame in result.Frames)
                {
                    token.ThrowIfCancellationRequested();
                    var reasons = new List<string>();
                    if (!string.IsNullOrEmpty(frame.Reason)) reasons.Add(frame.Reason);
                    bool ambiguousArc = reasons.Count != 0;
                    if (profileReason != null) reasons.Add(profileReason);
                    if (modelReason != null) reasons.Add(modelReason);
                    bool nominalCoordinates = !scene.SupportedCoordinates;
                    if (nominalCoordinates) reasons.Add(scene.NominalCoordinatesSupported
                        ? "Nominal coordinates only; couch pitch/roll or coordinate assurance is unconfirmed."
                        : "Supported coordinates are missing; no distance is available.");
                    bool coverage = (scene.Synthetic || scene.CtCoverageKnown) && !scene.ExternalTruncatedAtCtBoundary;
                    if (!coverage) reasons.Add(scene.ExternalTruncatedAtCtBoundary
                        ? "External reaches the CT boundary; anatomy coverage is truncated or uncertain."
                        : "CT extent coverage is unknown.");
                    if (!coverage && !string.IsNullOrWhiteSpace(scene.CtCoverageReason)) reasons.Add(scene.CtCoverageReason);
                    bool exactFrame = exact.ContainsKey(Key(frame.Pose.BeamId, frame.Pose.ControlPointIndex));
                    bool radialOnly = profile && scene.Profile.Kind == "RingBore";
                    reasons.Add(radialOnly ? "RingBore radial clearance only; finite-end/rim distance and safety margin are not established. No conservative lower bound or pass is available."
                        : profile && exactFrame ? "Exact sampled profile-envelope distance; pass applies only to supplied surfaces at this pose."
                        : profile ? "Bounded sampled profile distance; containment and exact minimum surface clearance are unresolved."
                        : "Nominal source no-fly model, not commissioned. Signed distance interval [lower bound, sampled upper bound]; not an exact minimum surface clearance.");

                    var body = new RoleResult(); var table = new RoleResult();
                    Func<BeamPoint3D, string, double> pointDistance = null;
                    if (modelReason == null && (scene.SupportedCoordinates || scene.NominalCoordinatesSupported))
                        pointDistance = profile ? CollisionSweepDistance.ProfilePointDistance(scene.Profile, frame.Pose)
                            : solid.AtPose(frame.Pose);
                    foreach (var surface in surfaces)
                    {
                        var role = surface.Surface.Role == "External" ? body : table;
                        role.Present = true;
                        if (!surface.Valid) { role.Complete = false; role.ModelComplete = false; role.Include(null, null); reasons.Add(surface.Reason); continue; }
                        if (!surface.Closed) reasons.Add(surface.Surface.Role + " mesh is open, nonmanifold or degenerate; it cannot pass.");
                        double? upper = null, lower = null; bool isExact = false;
                        List<CollisionFinding> findings;
                        if (exact.TryGetValue(Key(frame.Pose.BeamId, frame.Pose.ControlPointIndex), out findings))
                        {
                            var finding = findings.FirstOrDefault(f => f.SurfaceLabel == surface.Surface.Label);
                            if (finding != null)
                            {
                                upper = finding.ClearanceMm;
                                // Radial clearance after axial clipping can overstate the distance to a finite bore rim.
                                // It retains an overlap witness, but cannot establish a Euclidean margin or a pass.
                                if (!radialOnly) { lower = upper; isExact = upper.HasValue; }
                            }
                        }
                        else if (pointDistance != null)
                        {
                            var iso = frame.Pose.Isocenter;
                            var cacheKey = Tuple.Create(iso.X, iso.Y, iso.Z, surface.Surface.Label);
                            Tuple<double?, double?> cached;
                            if (stationaryBore && stationaryDistances.TryGetValue(cacheKey, out cached))
                            { lower = cached.Item1; upper = cached.Item2; }
                            else
                            {
                                if (spatialHead) CollisionSweepDistance.SampleSpatial(surface, pointDistance, token, spatialBudget, out lower, out upper);
                                else CollisionSweepDistance.Sample(surface, pointDistance, token, out lower, out upper);
                                if (stationaryBore) stationaryDistances.Add(cacheKey, Tuple.Create(lower, upper));
                            }
                        }
                        role.Include(lower, upper);
                        bool hit = upper.HasValue && upper.Value <= Epsilon && scene.SupportedCoordinates;
                        bool pass = isExact && lower.HasValue && lower.Value > scene.Profile.SafetyMarginMm &&
                            surface.Closed && scene.SupportedCoordinates && !ambiguousArc &&
                            (surface.Surface.Role != "External" || coverage);
                        role.Hit |= hit;
                        role.Complete &= pass;
                        // Separate evidence about captured surfaces from commissioning,
                        // CT completeness and setup assurance. Not a volume-containment
                        // or clinical clearance claim. Unknown/invalid parts cannot clear.
                        role.ModelHit |= upper.HasValue && upper.Value <= Epsilon;
                        role.ModelComplete &= spatialHead && surface.HasSurfaceFaces && !ambiguousArc &&
                            lower.HasValue && lower.Value > scene.SourceModel.WarningMarginMm;
                    }
                    if (!body.Present) reasons.Add("External surface is missing.");
                    if (!table.Present) reasons.Add("Support/table surface is missing.");
                    frame.BodyClearanceMm = body.Upper; frame.BodyLowerBoundMm = body.Lower;
                    frame.TableClearanceMm = table.Upper; frame.TableLowerBoundMm = table.Lower;
                    frame.BodyStatus = body.Status; frame.TableStatus = table.Status;
                    frame.ModelBodyStatus = spatialHead ? body.ModelStatus : null;
                    frame.ModelTableStatus = spatialHead ? table.ModelStatus : null;
                    frame.Status = body.Hit || table.Hit ? "hit" : body.Status == "pass" && table.Status == "pass" ? "pass" : "uncertain";
                    frame.Reason = string.Join(" ", reasons.Distinct());
                    if (stationaryBore && (scene.SupportedCoordinates || scene.NominalCoordinatesSupported))
                    {
                        CollisionRadialReview radial;
                        var radialKey = BoreKey(frame.Pose);
                        if (!radialReviews.TryGetValue(radialKey, out radial))
                        {
                            radial = EvaluateRadialBore(scene.SourceModel, frame.Pose, surfaces, token);
                            radialReviews.Add(radialKey, radial);
                        }
                        frame.RadialReview = radial;
                    }
                }
                token.ThrowIfCancellationRequested();
                result.Status = result.Frames.Any(f => f.Status == "hit") ? "hit" : result.Frames.All(f => f.Status == "pass") ? "pass" : "uncertain";
                result.Summary = (scene.Synthetic ? "SYNTHETIC demonstration. " : "") + result.Frames.Count + " sampled states: " +
                    result.Frames.Count(f => f.Status == "pass") + " pass, " + result.Frames.Count(f => f.Status == "uncertain") +
                    " uncertain, " + result.Frames.Count(f => f.Status == "hit") + " hit. " + Scope;
                return result;
            }
            catch (ArgumentException error) { return Uncertain(error.Message); }
        }

        private sealed class RoleResult
        {
            public bool Present, Hit, Complete = true, ModelHit, ModelComplete = true;
            public double? Upper, Lower;
            private bool missingBound;
            public string Status { get { return Hit ? "hit" : Present && Complete ? "pass" : "uncertain"; } }
            public string ModelStatus { get { return ModelHit ? "model-hit" : Present && ModelComplete ? "model-clear" : "uncertain"; } }
            public void Include(double? lower, double? upper)
            {
                if (upper.HasValue) Upper = Upper.HasValue ? Math.Min(Upper.Value, upper.Value) : upper;
                if (!lower.HasValue) { missingBound = true; Lower = null; }
                else if (!missingBound) Lower = Lower.HasValue ? Math.Min(Lower.Value, lower.Value) : lower;
            }
        }

        private static CollisionRadialReview EvaluateRadialBore(SourceCollisionModel model, CollisionPose pose,
            List<CollisionSweepDistance.PreparedSurface> surfaces, CancellationToken token)
        {
            var result = new CollisionRadialReview {
                Scope = "Open bore: radial distance to captured body/arms and table surfaces at nominal setup. No axial end wall. Clear means beyond the configured radial warning margin within this captured geometry only; anatomy outside CT, accessories and actual setup remain unverified. Not clinical clearance." };
            foreach (string role in new[] { "External", "Support" })
            {
                var selected = surfaces.Where(s => s.Surface.Role == role).ToList();
                double? clearance = null;
                if (selected.Count > 0 && selected.All(s => s.Valid && s.RadialSurfaceValid))
                {
                    double maxRadiusSquared = 0;
                    // Convex radial norm attains its maximum on a triangle at a vertex.
                    // Referenced vertices only; no end-cap or missing-anatomy inference.
                    foreach (var surface in selected)
                    {
                        var mesh = surface.Surface.Mesh;
                        for (int i = 0; i < mesh.TriangleIndices.Length; i++)
                        {
                            if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                            var p = mesh.Vertices[mesh.TriangleIndices[i]];
                            double x = p.X - pose.Isocenter.X, y = p.Y - pose.Isocenter.Y;
                            maxRadiusSquared = Math.Max(maxRadiusSquared, x * x + y * y);
                        }
                    }
                    clearance = model.BoreLimitRadiusMm - Math.Sqrt(maxRadiusSquared);
                }
                else result.Scope += " " + role + ": missing, invalid or degenerate surface geometry; radial clearance unavailable.";
                string status = !clearance.HasValue ? "uncertain" : clearance <= 0 ? "model-hit" :
                    clearance > model.BoreLimitRadiusMm - model.BoreWarningRadiusMm ? "clear" : "uncertain";
                if (role == "External") { result.BodyClearanceMm = clearance; result.BodyStatus = status; }
                else { result.TableClearanceMm = clearance; result.TableStatus = status; }
            }
            result.Status = result.BodyStatus == "model-hit" || result.TableStatus == "model-hit" ? "model-hit" :
                result.BodyStatus == "clear" && result.TableStatus == "clear" ? "clear" : "uncertain";
            return result;
        }

        private static List<CollisionSweepFrame> SampleFrames(List<CollisionPose> poses, CancellationToken token)
        {
            var output = new List<CollisionSweepFrame>();
            foreach (var beam in poses.GroupBy(p => p.BeamId, StringComparer.Ordinal))
            {
                var input = beam.OrderBy(p => p.ControlPointIndex).ToList();
                var selected = new List<CollisionSweepFrame>();
                AddFrame(selected, input[0], false, null);
                double unwrapped = input[0].GantryDegrees;
                for (int i = 1; i < input.Count; i++)
                {
                    token.ThrowIfCancellationRequested();
                    var start = input[i - 1]; var end = input[i];
                    double delta = Delta(end.GantryDegrees - start.GantryDegrees);
                    bool interpolable = Math.Abs(end.GantryDegrees - start.GantryDegrees) < 360 - Epsilon && Consistent(start, end, delta);
                    if (!interpolable)
                    {
                        const string reason = "Arc direction/source geometry is ambiguous or inconsistent; only exact captured endpoints are shown for this interval.";
                        AddFrame(selected, start, false, reason); AddFrame(selected, end, false, reason);
                        unwrapped += delta;
                    }
                    else if (Math.Abs(delta) > Epsilon)
                    {
                        double finish = unwrapped + delta;
                        double milestone = delta > 0 ? (Math.Floor((unwrapped + Epsilon) / 10) + 1) * 10
                            : (Math.Ceiling((unwrapped - Epsilon) / 10) - 1) * 10;
                        while (delta > 0 ? milestone <= finish + Epsilon : milestone >= finish - Epsilon)
                        {
                            double fraction = (milestone - unwrapped) / delta;
                            bool atEnd = Math.Abs(fraction - 1) < 1e-8;
                            AddFrame(selected, atEnd ? end : Interpolate(start, end, fraction, Normalize(milestone)), !atEnd, null,
                                fraction <= 0.5 ? start : end);
                            milestone += delta > 0 ? 10 : -10;
                        }
                        unwrapped = finish;
                    }
                    if (i == input.Count - 1) AddFrame(selected, end, false, null);
                    if (output.Count + selected.Count > MaximumFrames) throw new ArgumentException("Derived states exceed the bounded sweep budget.");
                }
                for (int i = 0; i < selected.Count; i++)
                { selected[i].Index = output.Count; selected[i].Pose.ControlPointIndex = i; output.Add(selected[i]); }
                if (output.Count > MaximumFrames) throw new ArgumentException("Derived states exceed the bounded sweep budget.");
            }
            return output;
        }

        private static void AddFrame(List<CollisionSweepFrame> frames, CollisionPose pose, bool interpolated, string reason, CollisionPose captured = null)
        {
            if (frames.Count > 0)
            {
                var previous = frames[frames.Count - 1];
                if (Math.Abs(previous.Pose.GantryDegrees - pose.GantryDegrees) < Epsilon &&
                    Math.Abs(Delta(previous.Pose.PatientSupportAngleDegrees - pose.PatientSupportAngleDegrees)) < Epsilon &&
                    Distance(previous.Pose.Isocenter, pose.Isocenter) < 1e-6 && Distance(previous.Pose.Source, pose.Source) < 1e-6)
                { if (reason != null) previous.Reason = reason; return; }
            }
            captured = captured ?? pose;
            frames.Add(new CollisionSweepFrame { Pose = Copy(pose), Interpolated = interpolated, Reason = reason,
                CapturedControlPointIndex = captured.ControlPointIndex, CapturedGantryDegrees = captured.GantryDegrees,
                CapturedCouchDegrees = captured.PatientSupportAngleDegrees });
        }

        private static bool Consistent(CollisionPose start, CollisionPose end, double delta)
        {
            if (Math.Abs(delta) > 45 || Distance(start.Isocenter, end.Isocenter) > 1e-3 ||
                Math.Abs(Delta(start.PatientSupportAngleDegrees - end.PatientSupportAngleDegrees)) > 1e-6) return false;
            var a = Direction(start); var b = Direction(end);
            double la = Length(a), lb = Length(b);
            if (Math.Abs(la - lb) > Math.Max(1e-3, la * 1e-5)) return false;
            double angle = Math.Acos(Math.Max(-1, Math.Min(1, Dot(a, b) / (la * lb)))) * 180 / Math.PI;
            if (Math.Abs(delta) > Epsilon && angle < 1e-6) return false;
            return Math.Abs(angle - Math.Abs(delta)) < 0.05;
        }

        private static CollisionPose Interpolate(CollisionPose start, CollisionPose end, double fraction, double degrees)
        {
            var a = Direction(start); var b = Direction(end); double la = Length(a), lb = Length(b);
            double omega = Math.Acos(Math.Max(-1, Math.Min(1, Dot(a, b) / (la * lb))));
            double left = Math.Sin((1 - fraction) * omega) / Math.Sin(omega) / la;
            double right = Math.Sin(fraction * omega) / Math.Sin(omega) / lb;
            double radius = la + fraction * (lb - la);
            return new CollisionPose { BeamId = start.BeamId, GantryDegrees = degrees, PatientSupportAngleDegrees = start.PatientSupportAngleDegrees, Isocenter = Copy(start.Isocenter),
                Source = new BeamPoint3D(start.Isocenter.X + radius * (left * a.X + right * b.X),
                    start.Isocenter.Y + radius * (left * a.Y + right * b.Y), start.Isocenter.Z + radius * (left * a.Z + right * b.Z)) };
        }

        private static void ValidateInputs(CollisionScene scene, CancellationToken token)
        {
            if (scene == null || scene.Poses == null || scene.Poses.Count == 0 || scene.Poses.Count > MaximumInputPoses ||
                scene.Surfaces == null || scene.Surfaces.Count > 32) throw new ArgumentException("Sweep poses or surfaces are missing or exceed bounded input limits.");
            var labels = new HashSet<string>(StringComparer.Ordinal);
            foreach (var s in scene.Surfaces)
                if (s == null || string.IsNullOrWhiteSpace(s.Label) || s.Label.Length > 256 || !labels.Add(s.Label) || (s.Role != "External" && s.Role != "Support"))
                    throw new ArgumentException("Surface roles and unique labels must be valid.");
            var identities = new HashSet<string>(StringComparer.Ordinal);
            foreach (var pose in scene.Poses)
            {
                token.ThrowIfCancellationRequested();
                if (pose == null || string.IsNullOrWhiteSpace(pose.BeamId) || pose.BeamId.Length > 256 || pose.ControlPointIndex < 0 ||
                    !Finite(pose.GantryDegrees) || pose.GantryDegrees < 0 || pose.GantryDegrees > 360 ||
                    !Finite(pose.PatientSupportAngleDegrees) || pose.PatientSupportAngleDegrees < 0 || pose.PatientSupportAngleDegrees > 360 ||
                    !ValidPoint(pose.Isocenter) || !ValidPoint(pose.Source) || Distance(pose.Isocenter, pose.Source) < 1e-6 ||
                    !identities.Add(Key(pose.BeamId, pose.ControlPointIndex))) throw new ArgumentException("A sweep pose is invalid, duplicated or lacks bounded source geometry.");
            }
        }
        private static CollisionSweepResult Uncertain(string reason) { return new CollisionSweepResult { Status = "uncertain", Summary = reason + " " + Scope }; }
        private static string Key(string beam, int index) { return beam + "\n" + index; }
        private static Tuple<double, double, double> BoreKey(CollisionPose pose)
        { return Tuple.Create(pose.Isocenter.X, pose.Isocenter.Y, pose.Isocenter.Z); }
        private static double Delta(double degrees) { return ((degrees + 540) % 360) - 180; }
        private static double Normalize(double degrees) { return ((degrees % 360) + 360) % 360; }
        private static CollisionPose Copy(CollisionPose p) { return new CollisionPose { BeamId = p.BeamId, ControlPointIndex = p.ControlPointIndex,
            GantryDegrees = p.GantryDegrees, PatientSupportAngleDegrees = p.PatientSupportAngleDegrees, Isocenter = Copy(p.Isocenter), Source = Copy(p.Source) }; }
        private static BeamPoint3D Copy(BeamPoint3D p) { return new BeamPoint3D(p.X, p.Y, p.Z); }
        private static BeamPoint3D Direction(CollisionPose p) { return new BeamPoint3D(p.Source.X - p.Isocenter.X, p.Source.Y - p.Isocenter.Y, p.Source.Z - p.Isocenter.Z); }
        private static double Dot(BeamPoint3D a, BeamPoint3D b) { return a.X * b.X + a.Y * b.Y + a.Z * b.Z; }
        private static double Length(BeamPoint3D a) { return Math.Sqrt(Dot(a, a)); }
        private static double Distance(BeamPoint3D a, BeamPoint3D b) { double x = a.X - b.X, y = a.Y - b.Y, z = a.Z - b.Z; return Math.Sqrt(x * x + y * y + z * z); }
        internal static bool ValidPoint(BeamPoint3D point) { return point != null && Finite(point.X) && Finite(point.Y) && Finite(point.Z) &&
            Math.Abs(point.X) <= 1000000 && Math.Abs(point.Y) <= 1000000 && Math.Abs(point.Z) <= 1000000; }
        internal static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
    }
}
