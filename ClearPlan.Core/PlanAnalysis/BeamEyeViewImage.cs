using Newtonsoft.Json;

namespace ClearPlan.Core.PlanAnalysis
{
    /// <summary>Detached in-memory HU volume. Origin is the center of voxel (0,0,0), axes are patient-space directions.</summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class CtVolume
    {
        public int SizeX { get; set; }
        public int SizeY { get; set; }
        public int SizeZ { get; set; }
        public double SpacingX { get; set; }
        public double SpacingY { get; set; }
        public double SpacingZ { get; set; }
        public BeamPoint3D Origin { get; set; }
        public BeamPoint3D XAxis { get; set; }
        public BeamPoint3D YAxis { get; set; }
        public BeamPoint3D ZAxis { get; set; }
        // Indexed (z * SizeY + y) * SizeX + x. Deliberately excluded from JSON.
        [JsonIgnore] public float[] HounsfieldUnits { get; set; }
    }

    /// <summary>In-memory DRR overview; pixels are never included in review JSON.</summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class BeamEyeViewImage
    {
        [JsonProperty] public int WidthPixels { get; set; }
        [JsonProperty] public int HeightPixels { get; set; }
        [JsonIgnore] public byte[] GrayscalePixels { get; set; }
        // Half extent in the isocenter plane; image spans [-ExtentMm,+ExtentMm].
        [JsonProperty] public double ExtentMm { get; set; }
        [JsonProperty] public bool Synthetic { get; set; }
        [JsonProperty] public string SourceStatus { get; set; }
        [JsonProperty] public string UnavailableReason { get; set; }
        [JsonProperty] public string ProjectionDescription { get; set; }
        [JsonProperty] public int ControlPointIndex { get; set; }
        [JsonProperty] public double GantryAngleDegrees { get; set; }
        [JsonProperty] public double CollimatorAngleDegrees { get; set; }
        // Calibrated BLD raster/aperture -> zero-collimator gantry display rotation, clockwise
        // in screen coordinates (Y down). Native capture derives this from paired outlines;
        // it is NOT inferred from CollimatorAngleDegrees. Null retains the labelled BLD frame.
        [JsonProperty] public double? BldToDisplayRotationDegrees { get; set; }
        [JsonProperty] public double PatientSupportAngleDegrees { get; set; }
    }
}
