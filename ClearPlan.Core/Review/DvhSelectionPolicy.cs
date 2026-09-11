using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Review
{
    public static class DvhSelectionPolicy
    {
        private static readonly Regex TargetAlias = new Regex(
            @"^\s*(?:EVAL[_. -]+)?(?<kind>PTV|CTV|GTV|ITV)(?=$|[0-9_ .]|-(?=[0-9]))",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex TargetExclusion = new Regex(
            @"(?:^|[^A-Z])(?:MINUS|WITHOUT|OHNE|EXCLUD(?:E|ED|ING))(?=$|[^A-Z])|[-−]\s*[A-Z]",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        public static string ClassifyTarget(string structureId, string dicomType)
        {
            string native = (dicomType ?? string.Empty).Trim().ToUpperInvariant();
            if (native == "EXTERNAL" || native == "SUPPORT") return string.Empty;
            if (TargetTokens.Contains(native)) return native;
            // Name-only inference is intentionally conservative: a leading target token
            // or explicit eval prefix, never an organ/ring/avoid prefix such as Brain-PTV.
            // Explicit exclusion wording and hyphen-to-name subtraction are not targets.
            // Native target types and exact manual/native selections remain authoritative.
            string id = (structureId ?? string.Empty).Trim();
            Match alias = TargetAlias.Match(id);
            if (!alias.Success || TargetExclusion.IsMatch(id.Substring(alias.Length))) return string.Empty;
            string kind = alias.Groups["kind"].Value.ToUpperInvariant();
            return kind;
        }

        /// <summary>
        /// Preserve every nonempty PTV, plus one largest verified contained CTV/GTV/ITV
        /// across all PTVs. Unknown geometry never supplies containment evidence.
        /// </summary>
        public static List<string> SelectTargetStructureIds(IEnumerable<TargetReviewStructure> structures)
        {
            var targets = (structures ?? Enumerable.Empty<TargetReviewStructure>())
                .Where(item => item != null && !item.IsEmpty && !string.IsNullOrWhiteSpace(item.StructureId))
                .ToList();
            var selected = targets.Where(item => ClassifyTarget(item.StructureId, item.DicomType) == "PTV")
                .Select(item => item.StructureId).Distinct(StringComparer.Ordinal).ToList();
            if (selected.Count == 0) return selected;
            var inner = targets.Where(item => new[] { "CTV", "GTV", "ITV" }
                .Contains(ClassifyTarget(item.StructureId, item.DicomType))).ToList();
            var largest = inner.Where(item => item.FullyContainedInPtv == true && ValidVolume(item.VolumeCc))
                .OrderByDescending(item => item.VolumeCc.Value)
                .ThenBy(item => item.StructureId, StringComparer.Ordinal).FirstOrDefault();
            // A larger unknown candidate prevents a claim that a smaller one is the largest.
            if (largest != null && !inner.Any(item => item.FullyContainedInPtv != false &&
                (!ValidVolume(item.VolumeCc) || (!item.FullyContainedInPtv.HasValue && item.VolumeCc >= largest.VolumeCc))))
                selected.Add(largest.StructureId);
            return selected;
        }

        private static bool ValidVolume(double? value)
        {
            return value.HasValue && value.Value > 0.0 && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value);
        }

        private static readonly string[] TargetTokens =
        {
            "PTV",
            "CTV",
            "GTV",
            "ITV"
        };

        /// <summary>Exact explicit requests (constraints or native/manual selections) override naming heuristics.</summary>
        public static bool ShouldSelect(
            string structureId,
            IEnumerable<string> requestedStructureIds)
        {
            string id = (structureId ?? string.Empty).Trim();
            if (id.Length == 0)
            {
                return false;
            }

            return (requestedStructureIds ??
                              Enumerable.Empty<string>())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Any(item => string.Equals(
                    item.Trim(),
                    id,
                    StringComparison.OrdinalIgnoreCase));
        }
    }

    /// <summary>Detached target facts; null containment means not established, not false.</summary>
    public sealed class TargetReviewStructure
    {
        public string StructureId { get; set; }
        public string DicomType { get; set; }
        public double? VolumeCc { get; set; }
        public bool IsEmpty { get; set; }
        public bool? FullyContainedInPtv { get; set; }
        public string ContainmentReason { get; set; }
    }
}
