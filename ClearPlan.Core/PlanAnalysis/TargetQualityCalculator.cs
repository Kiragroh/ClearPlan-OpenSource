using System;
using System.Collections.Generic;

namespace ClearPlan.Core.PlanAnalysis
{
    public sealed class ReviewTargetQuality
    {
        public string StructureId { get; set; }
        public string BodyStructureId { get; set; }
        public double? ReferenceDoseGy { get; set; }
        public double? TargetVolumeCm3 { get; set; }
        public double? CoveredTargetVolumeCm3 { get; set; }
        public double? BodyV100Cm3 { get; set; }
        public double? BodyV50Cm3 { get; set; }
        public double? D2Gy { get; set; }
        public double? D98Gy { get; set; }
        public double? PaddickCi { get; set; }
        public double? PlanCheckCi { get; set; }
        public double? GradientIndex { get; set; }
        public double? HomogeneityIndex { get; set; }
        public string Note { get; set; }

        // Display only: numerical availability is not clinical acceptance, and an empty note is not missing data.
        [Newtonsoft.Json.JsonIgnore]
        public string AvailabilityScope
        {
            get
            {
                int available = 0;
                foreach (double? value in new[] { PaddickCi, PlanCheckCi, GradientIndex, HomogeneityIndex })
                    if (value.HasValue && value.Value >= 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value)) available++;
                string status = available == 4 ? "Available" : available == 0 ? "Unavailable" : "Partly available";
                return string.IsNullOrWhiteSpace(Note) ? status : status + ". " + Note;
            }
        }
    }

    /// <summary>Detached DVH scalars only; cm3 and Gy. No TPS objects or modifications.</summary>
    public static class TargetQualityCalculator
    {
        public const string Definition = "Paddick CI = TV_Rx² / (TV × V_Rx); 1 / Paddick CI = (TV × V_Rx) / TV_Rx². " +
            "GI = V_50%Rx / V_Rx in the existing EXTERNAL (whole plan, not lesion-specific). " +
            "HI (PlanCheck) = (D2% - D98%) / Rx, not the D50-normalized variant. " +
            "Rx is the displayed plan prescription, not an inferred SIB target dose. Values are descriptive, not approval criteria.";

        public static ReviewTargetQuality Calculate(string structureId, double? referenceDoseGy, double? targetVolumeCm3,
            double? coveredTargetVolumeCm3, double? bodyV100Cm3, double? bodyV50Cm3, double? d2Gy, double? d98Gy)
        {
            var row = new ReviewTargetQuality { StructureId = structureId, ReferenceDoseGy = Valid(referenceDoseGy),
                TargetVolumeCm3 = Valid(targetVolumeCm3), CoveredTargetVolumeCm3 = Valid(coveredTargetVolumeCm3),
                BodyV100Cm3 = Valid(bodyV100Cm3), BodyV50Cm3 = Valid(bodyV50Cm3), D2Gy = Valid(d2Gy), D98Gy = Valid(d98Gy) };
            var notes = new List<string>();
            if (row.BodyV100Cm3 == 0) notes.Add("V100 = 0: no EXTERNAL volume receives the reference dose; CI/GI are undefined.");
            if (!(row.ReferenceDoseGy > 0))
            { row.Note = "Reference prescription dose unavailable; indices not evaluated."; return row; }
            if (row.TargetVolumeCm3 > 0 && row.BodyV100Cm3 > 0 && row.CoveredTargetVolumeCm3.HasValue &&
                row.CoveredTargetVolumeCm3 <= row.TargetVolumeCm3 && row.CoveredTargetVolumeCm3 <= row.BodyV100Cm3)
            {
                row.PaddickCi = Valid((row.CoveredTargetVolumeCm3 / row.TargetVolumeCm3) *
                    (row.CoveredTargetVolumeCm3 / row.BodyV100Cm3));
                if (row.PaddickCi > 0) row.PlanCheckCi = Valid(1 / row.PaddickCi);
                else notes.Add("Inverse CI undefined: no target volume receives Rx.");
            }
            else notes.Add("CI unavailable: target overlap/body DVH missing or inconsistent.");
            if (row.BodyV100Cm3 > 0 && row.BodyV50Cm3 >= row.BodyV100Cm3)
                row.GradientIndex = Valid(row.BodyV50Cm3 / row.BodyV100Cm3);
            else notes.Add("GI unavailable: body V50/V100 missing, zero or inconsistent.");
            if (row.D2Gy.HasValue && row.D98Gy.HasValue && row.D2Gy >= row.D98Gy)
                row.HomogeneityIndex = Valid((row.D2Gy - row.D98Gy) / row.ReferenceDoseGy);
            else notes.Add("HI unavailable: target D2/D98 missing or inconsistent.");
            row.Note = string.Join(" ", notes);
            return row;
        }

        public static bool HasSufficientCoverage(double? doseCoverage, double? samplingCoverage)
        {
            // ESAPI SamplingCoverage is not constrained to [0,1]; native voxel sampling
            // can slightly exceed 1. Match PlanCheck's lower bound, rejecting nonfinite data.
            return doseCoverage >= 0.9 && doseCoverage <= 1 && Valid(samplingCoverage).HasValue && samplingCoverage >= 0.9;
        }

        private static double? Valid(double? value)
        { return value.HasValue && value >= 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value) ? value : null; }
    }
}
