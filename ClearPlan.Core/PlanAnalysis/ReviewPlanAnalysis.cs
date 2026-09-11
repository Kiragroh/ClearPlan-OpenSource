using System.Collections.Generic;
using Newtonsoft.Json;

namespace ClearPlan.Core.PlanAnalysis
{
    public sealed class ReviewPlanAnalysis
    {
        public ReviewPlanAnalysis()
        {
            Beams = new List<ReviewBeamAnalysis>();
            TargetQuality = new List<ReviewTargetQuality>();
            Warnings = new List<string>();
            AvailableTargetStructureIds = new List<string>();
            PamTargetCandidates = new List<ReviewPamTargetCandidate>();
            PamStatus = "unavailable";
            PamReason = "Plan analysis has not been calculated.";
            PamWeightingMode = "MetersetMu";
            SmallApertureThresholdCm2 = 4;
        }

        public double? TotalMetersetMu { get; set; }
        public double? DosePerFractionGy { get; set; }
        public double? PlanNormalizationPercent { get; set; }
        // In-memory native freshness guard; not a hash of CT HU or structure contours.
        [JsonIgnore] public string NativePlanFingerprint { get; set; }
        public double? MuPerGy { get; set; }
        public double? Pam { get; set; }
        public double? MeanApertureAreaCm2 { get; set; }
        public double? SmallApertureFraction { get; set; }
        public double SmallApertureThresholdCm2 { get; set; }
        public string TargetStructureId { get; set; }
        public string TargetSelectionProvenance { get; set; }
        public string TargetSelectionMode { get; set; }
        public int PamTargetCandidateCount { get; set; }
        public int PamValidTargetCandidateCount { get; set; }
        public List<ReviewPamTargetCandidate> PamTargetCandidates { get; set; }
        // Native PlanCheck compatibility uses Beam.WeightFactor between beams.
        // Synthetic/import-only inputs explicitly retain meterset-MU aggregation.
        public string PamWeightingMode { get; set; }
        public List<string> AvailableTargetStructureIds { get; set; }
        public string PamStatus { get; set; }
        public string PamReason { get; set; }
        public string GeometryProvenance { get; set; }
        public string DoseRateProvenance { get; set; }
        public List<ReviewBeamAnalysis> Beams { get; set; }
        public List<ReviewTargetQuality> TargetQuality { get; set; }
        public string TargetQualityNote { get; set; }
        public List<string> Warnings { get; set; }
    }

    public sealed class ReviewPamTargetCandidate
    {
        public string TargetStructureId { get; set; }
        public double? Pam { get; set; }
        public string Status { get; set; }
        public string Reason { get; set; }
    }

    public sealed class ReviewBeamAnalysis
    {
        public ReviewBeamAnalysis()
        {
            ControlPoints = new List<ReviewControlPointSample>();
            GeometryStatus = "unavailable";
            DoseRateEstimateStatus = "Unavailable";
        }

        public string BeamId { get; set; }
        public int? BeamNumber { get; set; }
        public string MachineId { get; set; }
        public string MachineModelName { get; set; }
        public string MachineModel { get; set; }
        public string GantryDirection { get; set; }
        public string EnergyDisplay { get; set; }
        public string Technique { get; set; }
        public string MlcModel { get; set; }
        [JsonIgnore] public string NativeGeometryFingerprint { get; set; }
        public int MlcLayerCount { get; set; }
        public bool? HasJaws { get; set; }
        public double? MetersetMu { get; set; }
        public double? PamBeamWeightFactor { get; set; }
        public double? NominalDoseRateMuPerMin { get; set; }
        public string DoseRateEstimateStatus { get; set; }
        public string DoseRateEstimateReason { get; set; }
        public string DoseRateEstimateProfile { get; set; }
        public double? EstimatedBeamDurationSeconds { get; set; }
        public double? DoseRateEstimateMaxGantrySpeedDegreesPerSecond { get; set; }
        public double? Pam { get; set; }
        public double? MeanApertureAreaCm2 { get; set; }
        public double? SmallApertureFraction { get; set; }
        public double? MinimumApertureAreaCm2 { get; set; }
        public double? MaximumApertureAreaCm2 { get; set; }
        public string GeometryStatus { get; set; }
        public string GeometryReason { get; set; }
        public string PamReason { get; set; }
        public string MetersetReason { get; set; }
        public List<ReviewControlPointSample> ControlPoints { get; set; }
    }

    public sealed class ReviewControlPointSample
    {
        public ReviewControlPointSample()
        {
            TargetOutlines = new List<List<BeamPoint>>();
            TargetProjectionStrips = new List<ApertureRectangle>();
        }

        public int Index { get; set; }
        public double GantryAngleDegrees { get; set; }
        public double CollimatorAngleDegrees { get; set; }
        public double PatientSupportAngleDegrees { get; set; }
        public double CumulativeMetersetWeight { get; set; }
        // MU of the interval ending at this control point. First CP is zero.
        public double? IncrementalMetersetMu { get; set; }
        // Trapezoidal endpoint weight: half of each adjacent positive-MU interval.
        public double? MetricMetersetWeightMu { get; set; }
        public double? NominalDoseRateMuPerMin { get; set; }
        public double? PlannedDoseRateMuPerMin { get; set; }
        // PlanCheck-style interval estimate; neither a planned CP rate nor measured delivery.
        public double? EstimatedDoseRateMuPerMin { get; set; }
        public double? EstimatedSegmentDurationSeconds { get; set; }
        public double? ApertureAreaCm2 { get; set; }
        public double? TargetAreaCm2 { get; set; }
        public double? BlockedTargetFraction { get; set; }
        public string GeometryReason { get; set; }
        public string TargetProjectionReason { get; set; }
        public string TargetProjectionProvenance { get; set; }
        public double? TargetProjectionResolutionMm { get; set; }
        public double? NativeProjectionDifferenceFraction { get; set; }
        // Detached geometry is kept in memory for rendering, never in review JSON.
        [JsonIgnore]
        public ApertureGeometry Aperture { get; set; }
        [JsonIgnore]
        public List<List<BeamPoint>> TargetOutlines { get; set; }
        [JsonIgnore]
        public List<ApertureRectangle> TargetProjectionStrips { get; set; }
        [JsonIgnore]
        public double[] IsocenterMm { get; set; }
        [JsonIgnore]
        public BeamEyeViewImage BevImage { get; set; }
    }

    public sealed class BeamPoint
    {
        public BeamPoint() { }
        public BeamPoint(double x, double y) { X = x; Y = y; }
        public double X { get; set; }
        public double Y { get; set; }
    }

    public sealed class ApertureRectangle
    {
        public ApertureRectangle() { }
        public ApertureRectangle(double x1, double y1, double x2, double y2)
        { X1 = x1; Y1 = y1; X2 = x2; Y2 = y2; }
        public double X1 { get; set; }
        public double Y1 { get; set; }
        public double X2 { get; set; }
        public double Y2 { get; set; }
    }

    public sealed class ApertureLayer
    {
        public ApertureLayer() { LeafTravelAxis = "X"; }
        public string Label { get; set; }
        // X: positions along x, boundaries along y; Y: the converse.
        public string LeafTravelAxis { get; set; }
        public double[] LeafBoundariesMm { get; set; }
        public double[] Bank1PositionsMm { get; set; }
        public double[] Bank2PositionsMm { get; set; }
    }

    public sealed class ApertureGeometry
    {
        public ApertureGeometry()
        {
            Layers = new List<ApertureLayer>();
            EffectiveOpenings = new List<ApertureRectangle>();
        }
        // All distances are projected to isocenter in IEC beam-limiting-device mm.
        // Null jaws mean jawless; never synthesize jaws from zero/NaN coordinates.
        public ApertureRectangle Jaws { get; set; }
        // Fixed jawless field boundary (for example verified Halcyon virtual X/Y limits).
        // Clips the aperture but must never be labeled or counted as moving physical jaws.
        public ApertureRectangle FixedBoundingBox { get; set; }
        public List<ApertureLayer> Layers { get; set; }
        public List<ApertureRectangle> EffectiveOpenings { get; set; }
    }
}
