using System;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Review
{
    /// <summary>
    /// Defines the conservative automatic Ignore defaults used by the
    /// clinical PQM review. Valid numerical zero values remain visible.
    /// </summary>
    public static class PqmDefaultReviewPolicy
    {
        private static readonly Regex PartialHelperToken =
            new Regex(
                @"(?:^|[\s_-])(?:tl|tr|tb|teil)(?:$|[\s_-])",
                RegexOptions.Compiled |
                RegexOptions.CultureInvariant |
                RegexOptions.IgnoreCase);

        public static bool ShouldIgnore(
            string structureName,
            string achieved,
            string status)
        {
            string name = structureName ?? string.Empty;
            if (IsTarget(name))
            {
                return false;
            }

            if (PartialHelperToken.IsMatch(name))
            {
                return true;
            }

            return IsUnavailable(achieved) ||
                IsUnavailable(status);
        }

        private static bool IsTarget(string structureName)
        {
            return structureName.IndexOf(
                       "ptv",
                       StringComparison.OrdinalIgnoreCase) >= 0 ||
                structureName.IndexOf(
                    "ctv",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                structureName.IndexOf(
                    "gtv",
                    StringComparison.OrdinalIgnoreCase) >= 0 ||
                structureName.IndexOf(
                    "itv",
                    StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsUnavailable(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            return string.Equals(
                       normalized,
                       "Not evaluated",
                       StringComparison.OrdinalIgnoreCase) ||
                string.Equals(
                    normalized,
                    "Volume too small",
                    StringComparison.OrdinalIgnoreCase);
        }
    }
}
