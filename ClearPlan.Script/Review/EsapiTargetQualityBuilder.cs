using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.PlanAnalysis;
using ClearPlan.Core.Review;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace ClearPlan.Review
{
    /// <summary>Reads existing DVHs on the ESAPI owner thread. Never creates an isodose structure.</summary>
    public static class EsapiTargetQualityBuilder
    {
        public static void Populate(PlanSetup plan, ReviewPlanAnalysis analysis)
        {
            if (System.Windows.Application.Current != null) System.Windows.Application.Current.Dispatcher.VerifyAccess();
            analysis.TargetQuality = new List<ReviewTargetQuality>();
            analysis.TargetQualityNote = TargetQualityCalculator.Definition;
            if (plan == null || plan.StructureSet == null)
            { analysis.TargetQualityNote += " No structure set available."; return; }
            var targets = plan.StructureSet.Structures.Where(s => !s.IsEmpty && s.HasSegment &&
                DvhSelectionPolicy.ClassifyTarget(s.Id, s.DicomType) == "PTV").OrderBy(s => s.Id).ToList();
            analysis.TargetQualityNote += " Target recognition matches DVH: native PTV type or a leading PTV name (optionally eval-prefixed), without inferred organ/exclusion names; EXTERNAL/SUPPORT are excluded.";
            var bodies = plan.StructureSet.Structures.Where(s => !s.IsEmpty && s.HasSegment &&
                string.Equals(s.DicomType, "EXTERNAL", StringComparison.OrdinalIgnoreCase)).ToList();
            Structure body = bodies.Count == 1 ? bodies[0] : null;
            if (body == null) analysis.TargetQualityNote += " No unique EXTERNAL: CI/GI unavailable; target HI remains independent.";
            else analysis.TargetQualityNote += " Body scope: " + body.Id + ". Dose outside this contour is not included.";
            if (targets.Count > 1) analysis.TargetQualityNote += " Multiple PTVs: all rows use the plan Rx; verify target-specific prescription applicability (SIB/multiple lesions).";
            if (targets.Count == 0) analysis.TargetQualityNote += " No nonempty, segmented PTV recognized by the shared DVH target policy.";
            double? rx = Read(() => Gy(plan.TotalDose));
            bool hasDose = plan.Dose != null && plan.IsDoseValid;
            if (!hasDose) analysis.TargetQualityNote += " Current native dose missing or invalid: indices are not evaluated.";
            double? body100 = null, body50 = null;
            if (hasDose && rx > 0 && body != null && HasCoverage(plan, body))
            {
                body100 = Read(() => plan.GetVolumeAtDose(body, new DoseValue(rx.Value, DoseValue.DoseUnit.Gy), VolumePresentation.AbsoluteCm3));
                body50 = Read(() => plan.GetVolumeAtDose(body, new DoseValue(rx.Value * 0.5, DoseValue.DoseUnit.Gy), VolumePresentation.AbsoluteCm3));
            }
            foreach (var target in targets)
            {
                double? covered = null, d2 = null, d98 = null;
                if (hasDose && HasCoverage(plan, target))
                {
                    if (rx > 0) covered = Read(() => plan.GetVolumeAtDose(target, new DoseValue(rx.Value, DoseValue.DoseUnit.Gy), VolumePresentation.AbsoluteCm3));
                    d2 = Read(() => Gy(plan.GetDoseAtVolume(target, 2, VolumePresentation.Relative, DoseValuePresentation.Absolute)));
                    d98 = Read(() => Gy(plan.GetDoseAtVolume(target, 98, VolumePresentation.Relative, DoseValuePresentation.Absolute)));
                }
                var row = TargetQualityCalculator.Calculate(target.Id, rx, Read(() => target.Volume), covered, body100, body50, d2, d98);
                row.BodyStructureId = body == null ? null : body.Id;
                analysis.TargetQuality.Add(row);
            }
        }

        private static bool HasCoverage(PlanSetup plan, Structure structure)
        {
            try
            {
                double bin = plan.TotalDose.Unit == DoseValue.DoseUnit.cGy ? 10 : 0.1;
                var dvh = plan.GetDVHCumulativeData(structure, DoseValuePresentation.Absolute, VolumePresentation.AbsoluteCm3, bin);
                return dvh != null && TargetQualityCalculator.HasSufficientCoverage(dvh.Coverage, dvh.SamplingCoverage);
            }
            catch (Exception) { return false; }
        }
        private static double? Gy(DoseValue value)
        { return value.Unit == DoseValue.DoseUnit.Gy ? (double?)value.Dose : value.Unit == DoseValue.DoseUnit.cGy ? value.Dose / 100.0 : (double?)null; }
        private static double? Read(Func<double?> getter)
        {
            try { var value = getter(); return value.HasValue && value >= 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value) ? value : null; }
            catch (Exception) { return null; }
        }
    }
}
