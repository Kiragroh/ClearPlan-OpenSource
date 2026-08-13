using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.Review
{
    public static class DvhSelectionPolicy
    {
        private static readonly string[] ExcludedContours =
        {
            "BODY",
            "OUTER CONTOUR",
            "EXTERNAL",
            "KÖRPER"
        };

        private static readonly string[] TargetTokens =
        {
            "PTV",
            "CTV",
            "GTV",
            "ITV"
        };

        public static bool ShouldSelect(
            string structureId,
            IEnumerable<string> requestedStructureIds)
        {
            string id = (structureId ?? string.Empty).Trim();
            if (id.Length == 0 ||
                ExcludedContours.Any(item =>
                    string.Equals(item, id, StringComparison.OrdinalIgnoreCase)))
            {
                return false;
            }

            bool requested = (requestedStructureIds ??
                              Enumerable.Empty<string>())
                .Where(item => !string.IsNullOrWhiteSpace(item))
                .Any(item => string.Equals(
                    item.Trim(),
                    id,
                    StringComparison.OrdinalIgnoreCase));
            if (!requested)
            {
                return false;
            }

            string upper = id.ToUpperInvariant();
            bool target = TargetTokens.Any(upper.Contains);
            return target || !LooksLikePartialHelper(upper);
        }

        private static bool LooksLikePartialHelper(string upper)
        {
            if (upper.Contains("TEIL"))
            {
                return true;
            }

            return HasHelperToken(upper, "TL") ||
                   HasHelperToken(upper, "TB") ||
                   HasHelperToken(upper, "TR");
        }

        private static bool HasHelperToken(string value, string token)
        {
            return value.EndsWith("_" + token, StringComparison.Ordinal) ||
                   value.EndsWith("-" + token, StringComparison.Ordinal) ||
                   value.EndsWith(" " + token, StringComparison.Ordinal) ||
                   value.Contains("_" + token + "_") ||
                   value.Contains("-" + token + "-") ||
                   value.Contains(" " + token + " ");
        }
    }
}
