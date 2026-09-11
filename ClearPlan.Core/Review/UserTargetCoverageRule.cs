using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.Review
{
    /// <summary>One detached native prescription target; doses have explicit Gy units.</summary>
    public sealed class TargetPrescriptionDose
    {
        public string TargetId { get; set; }
        public string TargetType { get; set; }
        public double? DosePerFractionGy { get; set; }
        public int? FractionCount { get; set; }
    }

    /// <summary>
    /// The user's read-only ClearPlan rule, not a stock/literature constraint or an
    /// assigned Eclipse Clinical Goal. Never propagates Rx via geometric containment.
    /// </summary>
    public static class UserTargetCoverageRule
    {
        public const string SourceStableId = "source-user-target-coverage";
        public const string SourceLabel = "Default";

        public static ReviewPqmRow Evaluate(DefaultTargetReviewRule rule, string structureId,
            IEnumerable<TargetPrescriptionDose> prescriptionTargets, int? planFractionCount,
            double? d98Gy, bool? nativeDoseValid, string prescriptionUnavailableReason = null)
        {
            if (rule == null) throw new ArgumentNullException("rule");
            string id = structureId ?? string.Empty;
            string reason = prescriptionUnavailableReason;
            double? rxGy = null;
            var exact = (prescriptionTargets ?? Enumerable.Empty<TargetPrescriptionDose>())
                .Where(item => item != null && string.Equals(item.TargetId, id, StringComparison.Ordinal)).ToList();
            if (string.IsNullOrWhiteSpace(reason))
            {
                if (exact.Count != 1) reason = exact.Count == 0
                    ? "No exact native prescription target assignment; no plan-total or containing-PTV dose was substituted."
                    : "Multiple native prescription assignments for this target; dose assignment is ambiguous.";
                else if (!string.Equals(exact[0].TargetType, "Volume", StringComparison.OrdinalIgnoreCase))
                    reason = "Prescription is not assigned to a target volume.";
                else if (!planFractionCount.HasValue || planFractionCount <= 0 || !exact[0].FractionCount.HasValue ||
                    exact[0].FractionCount <= 0 || planFractionCount != exact[0].FractionCount)
                    reason = "Plan and target prescription fractions are missing or differ; partial-plan/course dose was not inferred.";
                else if (!ValidDose(exact[0].DosePerFractionGy) || exact[0].DosePerFractionGy <= 0)
                    reason = "Target prescription dose is missing, nonpositive, or lacks explicit Gy/cGy units.";
                else
                {
                    double total = exact[0].DosePerFractionGy.Value * exact[0].FractionCount.Value;
                    if (ValidDose(total)) rxGy = total;
                    else reason = "Target prescription total dose is not finite.";
                }
            }
            if (nativeDoseValid != true)
                reason = Join(reason, "Native plan dose validity is missing or false.");
            if (!ValidDose(d98Gy))
                reason = Join(reason, "Native " + rule.Metric + " is unavailable.");
            double? thresholdGy = rxGy * rule.PrescriptionPercent / 100.0;
            if (rxGy.HasValue && !ValidDose(thresholdGy))
            {
                thresholdGy = null;
                reason = Join(reason, "Configured threshold is not finite.");
            }
            bool evaluated = rxGy.HasValue && ValidDose(thresholdGy) && ValidDose(d98Gy) && nativeDoseValid == true && string.IsNullOrWhiteSpace(reason);
            bool passed = evaluated && Compare(d98Gy.Value, thresholdGy.Value, rule.Comparator);
            return new ReviewPqmRow
            {
                StableId = "default-" + rule.Id + "-" + id,
                TemplateCode = rule.Id,
                TemplateStructure = id,
                ResolvedStructureId = id,
                StructureOptions = new List<string> { id },
                Objective = rule.Metric,
                Comparator = rule.Comparator,
                Goal = thresholdGy,
                Variation = null,
                AchievedValue = nativeDoseValid == true && ValidDose(d98Gy) ? d98Gy : null,
                Unit = ReviewUnitCodes.Gray,
                Status = !evaluated ? ReviewStatusCodes.NotEvaluated : passed ? ReviewStatusCodes.Pass : ReviewStatusCodes.Fail,
                Severity = !evaluated ? ReviewSeverityCodes.Info : passed ? ReviewSeverityCodes.None : ReviewSeverityCodes.Error,
                SourceLabel = SourceLabel,
                MappingDescription = rxGy.HasValue ? "Exact native prescription target; matching fractions" : "Target Rx not resolved",
                Explanation = "Default rule from editable JSON: " + rule.Metric + " " + rule.Comparator + " " +
                    rule.PrescriptionPercent.ToString(System.Globalization.CultureInfo.InvariantCulture) + "% of exact target prescription. " +
                    (rule.Comparator == ">" || rule.Comparator == "<" ? "Strict comparison; equality fails. " : "") +
                    "Not a literature/stock constraint or an assigned Eclipse Clinical Goal. " +
                    (evaluated ? "Rx: native RTPrescription.Targets, exact TargetId, DosePerFraction x matching NumberOfFractions; achieved: native GetDoseAtVolume."
                        : "Not evaluated: " + reason)
            };
        }

        private static bool Compare(double actual, double threshold, string comparator)
        {
            switch (comparator)
            {
                case ">": return actual > threshold;
                case ">=": return actual >= threshold;
                case "<": return actual < threshold;
                case "<=": return actual <= threshold;
                case "=": return actual == threshold;
                default: throw new ArgumentException("Unsupported Default rule comparator.");
            }
        }

        private static string Join(string first, string second)
        {
            return string.IsNullOrWhiteSpace(first) ? second : first + " " + second;
        }

        private static bool ValidDose(double? value)
        {
            return value.HasValue && value.Value >= 0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value);
        }
    }
}
