using System.Collections.Generic;
using ClearPlan.Core.PlanAnalysis;

namespace ClearPlan.Core.Collision
{
    public sealed class CollisionScene
    {
        public CollisionScene() { Surfaces = new List<CollisionSurface>(); Poses = new List<CollisionPose>(); }
        public string PlanKey { get; set; }
        [Newtonsoft.Json.JsonIgnore] public string NativePlanFingerprint { get; set; }
        [Newtonsoft.Json.JsonIgnore] public Dictionary<string,string> NativeBeamFingerprints { get; set; }
        public bool Synthetic { get; set; }
        public string GeometryReason { get; set; }
        public bool SupportedCoordinates { get; set; }
        public List<CollisionSurface> Surfaces { get; set; }
        public List<CollisionPose> Poses { get; set; }
        public CollisionProfile Profile { get; set; }
        public string NativeMachineId { get; set; }
        public string PatientPosition { get; set; }
        public bool NominalCoordinatesSupported { get; set; }
        public bool CtCoverageKnown { get; set; }
        public bool ExternalTruncatedAtCtBoundary { get; set; }
        public string CtCoverageReason { get; set; }
        public SourceCollisionModel SourceModel { get; set; }
        public SourceCollisionScreeningResult SourceScreening { get; set; }
    }
    public sealed class CollisionSurface
    {
        public string Label { get; set; }
        public string Role { get; set; }
        public TargetMeshGeometry Mesh { get; set; }
    }
    public sealed class CollisionPose
    {
        public string BeamId { get; set; }
        public int ControlPointIndex { get; set; }
        public double GantryDegrees { get; set; }
        public double PatientSupportAngleDegrees { get; set; }
        public BeamPoint3D Isocenter { get; set; }
        public BeamPoint3D Source { get; set; }
    }
    public sealed class CollisionProfile
    {
        public string MachineId { get; set; }
        public string Kind { get; set; }
        public string Revision { get; set; }
        public string CommissionedBy { get; set; }
        public string Evidence { get; set; }
        public bool Commissioned { get; set; }
        public bool SyntheticOnly { get; set; }
        public double HeadCenterFromIsoMm { get; set; }
        public double HeadRadiusMm { get; set; }
        public double BoreRadiusMm { get; set; }
        public double BoreHalfLengthMm { get; set; }
        public double SafetyMarginMm { get; set; }
    }
    public sealed class CollisionResult
    {
        public CollisionResult() { Findings = new List<CollisionFinding>(); }
        public string Status { get; set; }
        public string Summary { get; set; }
        public List<CollisionFinding> Findings { get; set; }
    }
    public sealed class CollisionFinding
    {
        public string BeamId { get; set; }
        public int ControlPointIndex { get; set; }
        public double GantryDegrees { get; set; }
        public string SurfaceLabel { get; set; }
        public double? ClearanceMm { get; set; }
        public string Status { get; set; }
        public string Message { get; set; }
    }

    /// <summary>Detached adapter only. Source screening never supplies commissioning or coordinate assurance.</summary>
    public static class SourceCollisionSceneAdapter
    {
        public static void Attach(CollisionScene scene, SourceCollisionModelCatalog catalog, System.Threading.CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            if (scene == null) throw new System.ArgumentNullException(nameof(scene));
            scene.SourceModel = null; scene.SourceScreening = null;
            if (scene.Synthetic || catalog == null) return;
            scene.SourceModel = catalog.Select(scene.NativeMachineId);
            if (scene.SourceModel == null) return;
            var surfaces = new List<SourceCollisionSurface>();
            foreach (var surface in scene.Surfaces)
            {
                token.ThrowIfCancellationRequested();
                surfaces.Add(new SourceCollisionSurface { Label = surface.Label,
                    Role = surface.Role == "External" ? "external" : surface.Role == "Support" ? "support" : "unsupported",
                    Points = surface.Mesh == null ? null : surface.Mesh.Vertices });
            }
            scene.SourceScreening = SourceCollisionScreening.Evaluate(new SourceCollisionScreeningRequest {
                MachineId = scene.NativeMachineId, PatientPosition = scene.PatientPosition,
                CoordinateSystem = "DICOM-LPS-mm", NativeSourceCoordinates = scene.SupportedCoordinates,
                Surfaces = surfaces, Poses = scene.Poses }, catalog, token);
        }
    }
}
