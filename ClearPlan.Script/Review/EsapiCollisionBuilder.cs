using System;
using System.Linq;
using System.Threading;
using ClearPlan.Core.Collision;
using ClearPlan.Core.PlanAnalysis;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>Owner-thread, read-only primitive copies. No TPS modifications or guessed machine dimensions.</summary>
    public static class EsapiCollisionBuilder
    {
        public static CollisionScene Capture(PlanSetup plan, string key, CollisionProfileCatalog catalog, bool zeroPitchRollConfirmed, CancellationToken token, Action<string> diagnostic = null)
        {
            if (Thread.CurrentThread.GetApartmentState() != ApartmentState.STA)
                throw new InvalidOperationException("Native collision capture requires the ESAPI owner STA.");
            token.ThrowIfCancellationRequested();
            var scene = new CollisionScene { PlanKey = key, Synthetic = false, SupportedCoordinates = false };
            if (plan == null || plan.StructureSet == null) { scene.GeometryReason = "A native single plan with a structure set is required."; return scene; }
            var beams = plan.Beams.Where(b => !b.IsSetupField && !b.IsImagingTreatmentField).ToList();
            var machines = beams.Select(b => b.TreatmentUnit == null ? null : b.TreatmentUnit.Id).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            scene.NativeMachineId = machines.Count == 1 ? machines[0] : null;
            scene.PatientPosition = plan.TreatmentOrientation == PatientOrientation.HeadFirstSupine ? "HFS" : "unsupported";
            if (machines.Count == 1 && !string.IsNullOrWhiteSpace(machines[0]) && catalog != null)
                scene.Profile = catalog.Select(machines[0]);
            // This ESAPI exposes no table pitch/roll. Require separate, per-plan
            // operator assurance, never pretend these angles were read from ESAPI.
            bool coordinates = plan.TreatmentOrientation == PatientOrientation.HeadFirstSupine && machines.Count == 1;
            ControlPoint referenceTable = null;
            foreach (var structure in plan.StructureSet.Structures)
            {
                token.ThrowIfCancellationRequested();
                string role = string.Equals(structure.DicomType, "EXTERNAL", StringComparison.OrdinalIgnoreCase) ? "External" :
                    string.Equals(structure.DicomType, "SUPPORT", StringComparison.OrdinalIgnoreCase) ? "Support" : null;
                if (role == null || structure.IsEmpty) continue;
                var mesh = structure.MeshGeometry;
                if (diagnostic != null) diagnostic("collision-surface role=" + role + " vertices=" + (mesh == null ? 0 : mesh.Positions.Count) + " indices=" + (mesh == null ? 0 : mesh.TriangleIndices.Count));
                // Match the detached evaluator's bounded mesh capacity; retain every triangle.
                if (mesh == null || mesh.Positions.Count > 250000 || mesh.TriangleIndices.Count > 1500000)
                    throw new ArgumentException("Surface exceeds the bounded capture budget.");
                var vertices = mesh.Positions.Select(p => new BeamPoint3D(p.X, p.Y, p.Z)).ToList();
                if (vertices.Any(p => !Finite(p.X) || !Finite(p.Y) || !Finite(p.Z)))
                    throw new ArgumentException("Invalid surface coordinates.");
                if (mesh.TriangleIndices.Count % 3 != 0 || mesh.TriangleIndices.Any(index => index < 0 || index >= vertices.Count))
                    throw new ArgumentException("Invalid surface triangle indices.");
                // Neutral role label only: no structure IDs or personal information in diagnostics.
                scene.Surfaces.Add(new CollisionSurface { Role = role, Label = role + " " + (scene.Surfaces.Count(s => s.Role == role) + 1),
                    Mesh = new TargetMeshGeometry { Vertices = vertices, TriangleIndices = mesh.TriangleIndices.ToArray() } });
            }
            if (scene.Surfaces.Count(s => s.Role == "External") != 1) coordinates = false;
            // Capture the actual acquisition grid, never infer coverage from cropped report images.
            // A topologically closed mesh may contain artificial caps at the CT limits.
            scene.CtCoverageKnown = false;
            scene.ExternalTruncatedAtCtBoundary = false;
            var image = plan.StructureSet.Image;
            if (image != null)
                CollisionCtCoverage.Apply(scene, new BeamPoint3D(image.Origin.x, image.Origin.y, image.Origin.z),
                    new[] { new BeamPoint3D(image.XDirection.x, image.XDirection.y, image.XDirection.z),
                        new BeamPoint3D(image.YDirection.x, image.YDirection.y, image.YDirection.z),
                        new BeamPoint3D(image.ZDirection.x, image.ZDirection.y, image.ZDirection.z) },
                    new[] { image.XRes, image.YRes, image.ZRes }, new[] { image.XSize, image.YSize, image.ZSize });
            else scene.CtCoverageReason = "Planning CT is missing; body coverage is unknown.";
            // Pair the detached MLC only with the same captured native plan/beam state.
            // A missing stamp withholds the inset, never the inspectable body surfaces.
            scene.NativeBeamFingerprints = new System.Collections.Generic.Dictionary<string,string>(StringComparer.Ordinal);
            try { scene.NativePlanFingerprint = EsapiNativeGeometryFingerprint.CapturePlan(plan,token); }
            catch (OperationCanceledException) { throw; }
            catch (Exception) { scene.NativePlanFingerprint = null; }
            foreach (var beam in beams)
            {
                token.ThrowIfCancellationRequested();
                try { scene.NativeBeamFingerprints[beam.Id] = EsapiNativeGeometryFingerprint.CaptureBeam(beam,token); }
                catch (OperationCanceledException) { throw; }
                catch (Exception) { scene.NativeBeamFingerprints[beam.Id] = null; }
                var points = beam.ControlPoints.ToList();
                if (diagnostic != null) diagnostic("collision-field ordinal=" + (beams.IndexOf(beam) + 1) + " controlPoints=" + points.Count);
                if (points.Count == 0 || scene.Poses.Count + points.Count > CollisionSweep.MaximumInputPoses) { coordinates = false; break; }
                if (referenceTable == null) referenceTable = points[0];
                if (diagnostic != null)
                {
                    var p = beam.GetSourceLocation(90); var iso = beam.IsocenterPosition;
                    diagnostic(string.Format(System.Globalization.CultureInfo.InvariantCulture,
                        "collision-field-frame couch={0:0.###} source90-relative-mm={1:0.###},{2:0.###},{3:0.###} table={4},{5},{6}",
                        points[0].PatientSupportAngle,p.x-iso.x,p.y-iso.y,p.z-iso.z,
                        points[0].TableTopLateralPosition,points[0].TableTopLongitudinalPosition,points[0].TableTopVerticalPosition));
                }
                // Static ESAPI endpoints may both expose Index == -1. Use their ordered
                // positions, as in the detached BEV/plan-analysis snapshot, for pose identity.
                for (int pointIndex = 0; pointIndex < points.Count; pointIndex++)
                {
                    var cp = points[pointIndex];
                    token.ThrowIfCancellationRequested();
                    if (!Finite(cp.PatientSupportAngle) || Math.Abs(Math.IEEERemainder(cp.PatientSupportAngle - points[0].PatientSupportAngle, 360)) > 1e-6 ||
                        !Same(referenceTable.TableTopLateralPosition, cp.TableTopLateralPosition) ||
                        !Same(referenceTable.TableTopLongitudinalPosition, cp.TableTopLongitudinalPosition) ||
                        !Same(referenceTable.TableTopVerticalPosition, cp.TableTopVerticalPosition)) coordinates = false;
                    var source = beam.GetSourceLocation(cp.GantryAngle);
                    var iso = beam.IsocenterPosition;
                    if (!Finite(cp.GantryAngle) || !Finite(source.x) || !Finite(source.y) || !Finite(source.z) ||
                        !Finite(iso.x) || !Finite(iso.y) || !Finite(iso.z)) throw new ArgumentException("Invalid source or isocenter coordinates.");
                    scene.Poses.Add(new CollisionPose { BeamId = beam.Id, ControlPointIndex = pointIndex, GantryDegrees = cp.GantryAngle,
                        PatientSupportAngleDegrees = cp.PatientSupportAngle,
                        Isocenter = new BeamPoint3D(iso.x, iso.y, iso.z), Source = new BeamPoint3D(source.x, source.y, source.z) });
                }
            }
            scene.NominalCoordinatesSupported = coordinates;
            // Noncoplanar display uses the beam's native source positions in the
            // patient frame. It is nominal only until separate local acceptance.
            scene.SupportedCoordinates = coordinates && zeroPitchRollConfirmed &&
                scene.Poses.All(p => Math.Abs(Math.IEEERemainder(p.PatientSupportAngleDegrees,360)) < 1e-6);
            scene.GeometryReason = scene.SupportedCoordinates
                ? "Native HFS surfaces/source positions; fixed zero couch yaw and fixed translation across fields. Fixed zero pitch/roll is operator-confirmed, NOT supplied by ESAPI. Actual treatment-day 6D corrections are not captured. Surface coverage/accessories require local verification. Sampled positions only."
                : coordinates ? "Nominal HFS native source positions; couch yaw is fixed within each field and may differ between fields. Fixed table translation. No patient-mesh rotation is guessed. Noncoplanar geometry, pitch/roll and actual treatment setup are not clinically validated."
                : "Unsupported or ambiguous geometry: requires HFS, one External, one machine, fixed yaw within each field and fixed table translation across fields. No coordinate transform inferred.";
            return scene;
        }
        private static bool Same(double a, double b)
        {
            // Unknown table coordinates are not evidence of a fixed treatment position.
            return Finite(a) && Finite(b) && Math.Abs(a - b) < 1e-6;
        }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }
    }
}
