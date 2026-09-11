using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.PlanAnalysis
{
    /// <summary>Read-only target policy. Inputs contain only nonempty, segmented structures.
    /// PTV eligibility uses the shared DVH target classification, not an arbitrary name prefix.</summary>
    public static class PamTargetSelection
    {
        public static string Resolve(string requestedId, string planTargetId,
            IEnumerable<string> eligibleStructureIds, IEnumerable<string> eligiblePtvIds,
            out string provenance, out string reason)
        {
            provenance = reason = null;
            var eligible = (eligibleStructureIds ?? Enumerable.Empty<string>())
                .Where(id => !string.IsNullOrWhiteSpace(id)).ToList();
            if (!string.IsNullOrWhiteSpace(requestedId))
            {
                string selected = UniqueMatch(eligible, requestedId.Trim());
                if (selected == null)
                    reason = "The explicit target is missing, ambiguous, empty, or not segmented. No fallback replaces an explicit choice.";
                else provenance = "Explicit target selection; plan data unchanged.";
                return selected;
            }
            var ptvs = AutomaticCandidateIds(eligible, eligiblePtvIds);
            if (ptvs.Count > 1)
            {
                reason = "Multiple eligible PTVs (" + ptvs.Count + "). Automatic PAM calculation compares every eligible PTV and selects the lowest complete available PAM; no TPS target assignment is changed.";
                return null;
            }
            if (!string.IsNullOrWhiteSpace(planTargetId))
            {
                string selected = UniqueMatch(eligible, planTargetId.Trim());
                if (selected != null)
                { provenance = "Valid PlanSetup.TargetVolumeID; plan data unchanged."; return selected; }
            }
            if (ptvs.Count == 1)
            {
                string selected = UniqueMatch(eligible, ptvs[0]);
                if (selected != null)
                {
                    provenance = "Automatically selected the unique nonempty, segmented PTV recognized by the shared DVH target policy because PlanSetup.TargetVolumeID is " +
                        (string.IsNullOrWhiteSpace(planTargetId) ? "unset" : "invalid") + "; no TPS target assignment was changed.";
                    return selected;
                }
            }
            reason = "No unique eligible nonempty, segmented PTV recognized by the shared DVH target policy and no valid plan target. Select a PAM target explicitly.";
            return null;
        }

        public static List<string> AutomaticCandidateIds(IEnumerable<string> eligibleStructureIds,
            IEnumerable<string> eligiblePtvIds)
        {
            var eligible = (eligibleStructureIds ?? Enumerable.Empty<string>())
                .Where(id => !string.IsNullOrWhiteSpace(id)).ToList();
            return (eligiblePtvIds ?? Enumerable.Empty<string>())
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Select(id => UniqueMatch(eligible, id)).Where(id => id != null)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(id => id, StringComparer.OrdinalIgnoreCase).ThenBy(id => id, StringComparer.Ordinal).ToList();
        }

        /// <summary>Compare completed detached analyses only. Null/nonfinite/partial results
        /// are excluded, never interpreted as zero. This does not assess clinical superiority.</summary>
        public static ReviewPlanAnalysis SelectLowestAvailable(IEnumerable<ReviewPlanAnalysis> candidates,
            ReviewPlanAnalysis fallback)
        {
            var all = (candidates ?? Enumerable.Empty<ReviewPlanAnalysis>()).Where(p => p != null).ToList();
            var valid = all.Where(CompletePam)
                .OrderBy(p => p.Pam.Value).ThenBy(p => p.TargetStructureId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(p => p.TargetStructureId, StringComparer.Ordinal).ToList();
            var selected = valid.FirstOrDefault() ?? PlanAnalysisSnapshot.Copy(fallback ?? new ReviewPlanAnalysis());
            selected.TargetSelectionMode = "AutomaticLowestPam";
            selected.PamTargetCandidateCount = all.Count;
            selected.PamValidTargetCandidateCount = valid.Count;
            selected.PamTargetCandidates = all.OrderBy(p => p.TargetStructureId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(p => p.TargetStructureId, StringComparer.Ordinal)
                .Select(p => new ReviewPamTargetCandidate { TargetStructureId = p.TargetStructureId,
                    Pam = CompletePam(p) ? p.Pam : null, Status = CompletePam(p) ? "available" : "unavailable",
                    Reason = CompletePam(p) ? null : p.PamReason }).ToList();
            selected.TargetSelectionProvenance = "Automatic lowest complete available PAM among " + valid.Count + " of " + all.Count +
                " eligible PTV candidates. Equal PAM values use ordinal target-ID order. This numerical minimum is not a claim of clinical superiority; no TPS target assignment was changed.";
            if (valid.Count > 0)
                selected.PamReason = (selected.PamWeightingMode == "BeamWeightFactor" ? PlanAnalysisCalculator.NativePamDefinition : PlanAnalysisCalculator.PamDefinition) +
                    " Target selection: " + selected.TargetSelectionProvenance;
            else
            {
                selected.TargetStructureId = null;
                foreach (var beam in selected.Beams) foreach (var cp in beam.ControlPoints)
                {
                    cp.TargetOutlines.Clear();
                    cp.TargetProjectionStrips.Clear();
                    cp.TargetProjectionProvenance = null;
                    cp.TargetProjectionResolutionMm = null;
                    cp.NativeProjectionDifferenceFraction = null;
                    cp.TargetProjectionReason = "No eligible PTV has a complete finite available PAM.";
                }
                PlanAnalysisCalculator.Calculate(selected);
                selected.PamReason = "No eligible PTV has a complete finite available PAM. " + selected.TargetSelectionProvenance;
            }
            return selected;
        }

        private static bool CompletePam(ReviewPlanAnalysis plan)
        {
            if (plan.PamStatus != "available" || string.IsNullOrWhiteSpace(plan.TargetStructureId) || !ValidPam(plan.Pam) ||
                (plan.PamWeightingMode != "MetersetMu" && plan.PamWeightingMode != "BeamWeightFactor") ||
                !plan.TotalMetersetMu.HasValue || !Finite(plan.TotalMetersetMu.Value) || plan.TotalMetersetMu <= 0 ||
                plan.Beams == null || plan.Beams.Count == 0 || plan.Beams.Any(b => b == null ||
                    !b.MetersetMu.HasValue || !Finite(b.MetersetMu.Value) || b.MetersetMu < 0)) return false;
            if (plan.PamWeightingMode == "BeamWeightFactor" && plan.Beams.Any(b => !b.PamBeamWeightFactor.HasValue ||
                !Finite(b.PamBeamWeightFactor.Value) || b.PamBeamWeightFactor < 0)) return false;
            var active = plan.Beams.Where(b => plan.PamWeightingMode == "BeamWeightFactor" ? b.PamBeamWeightFactor > 0 : b.MetersetMu > 0).ToList();
            return active.Count > 0 && active.All(b => ValidPam(b.Pam) && b.GeometryStatus == "available" &&
                b.ControlPoints != null && b.ControlPoints.Count >= 2 &&
                b.ControlPoints.Any(cp => cp.MetricMetersetWeightMu > 0) &&
                b.ControlPoints.All(cp => cp.MetricMetersetWeightMu.HasValue && Finite(cp.MetricMetersetWeightMu.Value) &&
                    cp.MetricMetersetWeightMu >= 0) &&
                b.ControlPoints.Where(cp => cp.MetricMetersetWeightMu > 0).All(cp => ValidPam(cp.BlockedTargetFraction) &&
                    cp.TargetAreaCm2.HasValue && Finite(cp.TargetAreaCm2.Value) && cp.TargetAreaCm2 > 0));
        }

        private static bool ValidPam(double? value)
        { return value.HasValue && Finite(value.Value) && value >= 0 && value <= 1; }
        private static bool Finite(double value) { return !double.IsNaN(value) && !double.IsInfinity(value); }

        private static string UniqueMatch(List<string> ids, string requested)
        {
            var matches = ids.Where(id => string.Equals(id, requested, StringComparison.OrdinalIgnoreCase)).ToList();
            return matches.Count == 1 ? matches[0] : null;
        }
    }
}
