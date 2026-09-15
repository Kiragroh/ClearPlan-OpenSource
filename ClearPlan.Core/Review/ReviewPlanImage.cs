using System.Collections.Generic;
using Newtonsoft.Json;

namespace ClearPlan.Core.Review
{
    /// <summary>Detached, windowed overview pixels. Never retains an ESAPI image or patient identifier.</summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewPlanImage
    {
        public ReviewPlanImage() { Overlays = new List<ReviewImageOverlay>(); }
        [JsonProperty] public string PlanKey { get; set; }
        [JsonProperty] public string Kind { get; set; }
        [JsonProperty] public string Title { get; set; }
        [JsonProperty] public string Caption { get; set; }
        [JsonProperty] public string SourceStatus { get; set; }
        [JsonProperty] public string UnavailableReason { get; set; }
        [JsonProperty] public bool Synthetic { get; set; }
        [JsonProperty] public int WidthPixels { get; set; }
        [JsonProperty] public int HeightPixels { get; set; }
        [JsonProperty] public byte[] GrayscalePixels { get; set; }
        [JsonProperty] public double PixelSpacingXMillimeters { get; set; }
        [JsonProperty] public double PixelSpacingYMillimeters { get; set; }
        [JsonProperty] public string LeftOrientation { get; set; }
        [JsonProperty] public string RightOrientation { get; set; }
        [JsonProperty] public string TopOrientation { get; set; }
        [JsonProperty] public string BottomOrientation { get; set; }
        [JsonProperty] public double? IsocenterPixelX { get; set; }
        [JsonProperty] public double? IsocenterPixelY { get; set; }
        [JsonProperty] public List<ReviewImageOverlay> Overlays { get; set; }
        [JsonProperty] public string OverlaySummary { get; set; }
        [JsonProperty] public ReviewImageDoseRegion DoseFocusRegion { get; set; }
        [JsonProperty] public ReviewImageDosePlane DosePlane { get; set; }
    }

    /// <summary>Native interpolated Gy on the CT plane; NaN marks unavailable coverage. No vendor objects.</summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewImageDosePlane
    {
        [JsonProperty] public int Columns { get; set; }
        [JsonProperty] public int Rows { get; set; }
        [JsonProperty] public double PrescriptionGy { get; set; }
        [JsonProperty] public double[] SamplesGy { get; set; }
        [JsonProperty] public string Source { get; set; }
        [JsonIgnore] public bool IsValid { get { return Columns >= 2 && Rows >= 2 && Columns <= 512 && Rows <= 512 &&
            SamplesGy != null && SamplesGy.Length == Columns * Rows && IsodoseDisplayConfiguration.Finite(PrescriptionGy) && PrescriptionGy > 0 &&
            System.Linq.Enumerable.All(SamplesGy, v => double.IsNaN(v) || (IsodoseDisplayConfiguration.Finite(v) && v >= 0)); } }
    }

    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewImageDoseRegion
    {
        [JsonProperty] public double PrescriptionPercent { get; set; }
        [JsonProperty] public double ThresholdGy { get; set; }
        [JsonProperty] public double MinPixelX { get; set; }
        [JsonProperty] public double MaxPixelX { get; set; }
        [JsonProperty] public double MinPixelY { get; set; }
        [JsonProperty] public double MaxPixelY { get; set; }
        [JsonProperty] public bool CoverageLimited { get; set; }
    }

    /// <summary>Detached native-source annotation; coordinates refer to CT pixel centers, origin top left.</summary>
    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewImageOverlay
    {
        public ReviewImageOverlay() { Paths = new List<ReviewImagePath>(); }
        [JsonProperty] public string Kind { get; set; }
        [JsonProperty] public string Label { get; set; }
        [JsonProperty] public string ColorHex { get; set; }
        [JsonProperty] public string Source { get; set; }
        [JsonProperty] public string SourceStatus { get; set; }
        [JsonProperty] public string UnavailableReason { get; set; }
        [JsonProperty] public double? DoseGy { get; set; }
        [JsonProperty] public double? PrescriptionPercent { get; set; }
        [JsonProperty] public List<ReviewImagePath> Paths { get; set; }
    }

    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewImagePath
    {
        public ReviewImagePath() { Points = new List<ReviewImagePoint>(); }
        [JsonProperty] public bool Closed { get; set; }
        [JsonProperty] public List<ReviewImagePoint> Points { get; set; }
    }

    [JsonObject(MemberSerialization.OptIn)]
    public sealed class ReviewImagePoint
    {
        [JsonProperty] public double X { get; set; }
        [JsonProperty] public double Y { get; set; }
    }
}
