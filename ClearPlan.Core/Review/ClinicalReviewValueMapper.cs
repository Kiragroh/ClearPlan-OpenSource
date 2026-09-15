using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Review
{
    public sealed class ReviewStatusAndSeverity
    {
        public ReviewStatusAndSeverity(string status, string severity)
        {
            Status = status ?? ReviewStatusCodes.NotEvaluated;
            Severity = severity ?? ReviewSeverityCodes.Info;
        }

        public string Status { get; private set; }

        public string Severity { get; private set; }
    }

    public static class ClinicalReviewValueMapper
    {
        private static readonly Regex NumericValuePattern =
            new Regex(
                @"[-+]?\d+(?:[.,]\d+)?",
                RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex DicomUidPattern =
            new Regex(
                @"(?<![\d.])\d+(?:\.\d+){4,}(?![\d.])",
                RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private static readonly Regex AbsolutePathPattern =
            new Regex(
                @"(?<![A-Za-z0-9])(?:[A-Za-z]:[\\/]|\\\\|/)",
                RegexOptions.CultureInvariant | RegexOptions.Compiled);

        public static ReviewStatusAndSeverity MapPlanCheckStatus(string value)
        {
            string normalized = Normalize(value);
            if (normalized.StartsWith("3", StringComparison.Ordinal) ||
                normalized == "ok")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Pass,
                    ReviewSeverityCodes.None);
            }

            if (normalized.StartsWith("2", StringComparison.Ordinal) ||
                normalized == "variation")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Variation,
                    ReviewSeverityCodes.Warning);
            }

            if (normalized.StartsWith("1", StringComparison.Ordinal) ||
                normalized == "deviation" ||
                normalized == "error")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Fail,
                    ReviewSeverityCodes.Error);
            }

            if (normalized.StartsWith("0", StringComparison.Ordinal) ||
                normalized == "information" ||
                normalized == "info")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Info,
                    ReviewSeverityCodes.Info);
            }

            return NotEvaluated();
        }

        public static ReviewStatusAndSeverity MapPqmStatus(string value)
        {
            string normalized = Normalize(value);
            if (normalized == "goal")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Pass,
                    ReviewSeverityCodes.None);
            }

            if (normalized == "variation")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Variation,
                    ReviewSeverityCodes.Warning);
            }

            if (normalized == "not met")
            {
                return new ReviewStatusAndSeverity(
                    ReviewStatusCodes.Fail,
                    ReviewSeverityCodes.Error);
            }

            return NotEvaluated();
        }

        public static string MapFieldStatus(string currentValue, string expectedValue)
        {
            if (string.IsNullOrWhiteSpace(currentValue) ||
                string.IsNullOrWhiteSpace(expectedValue))
            {
                return ReviewStatusCodes.NotEvaluated;
            }

            return string.Equals(
                    currentValue.Trim(),
                    expectedValue.Trim(),
                    StringComparison.Ordinal)
                ? ReviewStatusCodes.Pass
                : ReviewStatusCodes.Variation;
        }

        public static bool TryParseClinicalDouble(
            string value,
            out double parsedValue)
        {
            parsedValue = 0.0;
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            Match match = NumericValuePattern.Match(value);
            if (!match.Success)
            {
                return false;
            }

            string invariantText = match.Value.Replace(',', '.');
            if (!double.TryParse(
                    invariantText,
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out parsedValue))
            {
                parsedValue = 0.0;
                return false;
            }

            if (double.IsNaN(parsedValue) || double.IsInfinity(parsedValue))
            {
                parsedValue = 0.0;
                return false;
            }

            return true;
        }

        /// <summary>
        /// Captures an optional absolute-dose display value. Missing native values
        /// remain null; this does not establish dose validity or goal conformance.
        /// </summary>
        public static double? ReadOptionalDoseInGray(Func<double> readDoseInGray)
        {
            if (readDoseInGray == null) throw new ArgumentNullException("readDoseInGray");
            try
            {
                double value = readDoseInGray();
                return double.IsNaN(value) || double.IsInfinity(value) || value < 0.0
                    ? (double?)null : value;
            }
            catch (Exception)
            {
                // Optional ESAPI properties can be unavailable independently.
                // Never substitute zero or infer total dose from prescription.
                return null;
            }
        }

        public static double ConvertDoseToGray(double value, string doseUnit)
        {
            if (double.IsNaN(value) || double.IsInfinity(value))
            {
                throw new ArgumentOutOfRangeException(
                    "value",
                    "Dose must be finite.");
            }

            if (string.Equals(
                    doseUnit,
                    ReviewUnitCodes.Gray,
                    StringComparison.OrdinalIgnoreCase))
            {
                return value;
            }

            if (string.Equals(
                    doseUnit,
                    "cGy",
                    StringComparison.OrdinalIgnoreCase))
            {
                return value / 100.0;
            }

            throw new ArgumentException(
                "Only Gy and cGy dose units are supported.",
                "doseUnit");
        }

        public static string CreateUniqueStableId(
            string candidate,
            string fallback,
            ISet<string> usedIds)
        {
            if (usedIds == null)
            {
                throw new ArgumentNullException("usedIds");
            }

            string safeCandidate = SanitizeClinicalLabel(candidate, fallback);
            string baseId = ToStableId(safeCandidate);
            if (string.IsNullOrWhiteSpace(baseId))
            {
                baseId = ToStableId(
                    SanitizeClinicalLabel(fallback, "clinical-item"));
            }

            if (string.IsNullOrWhiteSpace(baseId))
            {
                baseId = "clinical-item";
            }

            string uniqueId = baseId;
            int occurrence = 2;
            while (!usedIds.Add(uniqueId))
            {
                uniqueId = string.Format(
                    CultureInfo.InvariantCulture,
                    "{0}-{1}",
                    baseId,
                    occurrence);
                occurrence++;
            }

            return uniqueId;
        }

        public static string SanitizeClinicalLabel(
            string value,
            string fallback)
        {
            string safeFallback = IsUnsafe(fallback)
                ? "Clinical item"
                : (fallback ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(safeFallback))
            {
                safeFallback = "Clinical item";
            }

            if (string.IsNullOrWhiteSpace(value) || IsUnsafe(value))
            {
                return safeFallback;
            }

            return value.Trim();
        }

        public static string ExtractComparator(string value)
        {
            string trimmed = (value ?? string.Empty).TrimStart();
            foreach (string comparator in new[] { "<=", ">=", "<", ">", "=" })
            {
                if (trimmed.StartsWith(comparator, StringComparison.Ordinal))
                {
                    return comparator;
                }
            }

            return string.Empty;
        }

        public static string InferUnit(string value)
        {
            var outputUnit = Regex.Match(value ?? string.Empty, @"\[(?<unit>Gy|cGy|%|cc|cm3)\]\s*$", RegexOptions.IgnoreCase);
            if (outputUnit.Success) value = outputUnit.Groups["unit"].Value;
            string normalized = Normalize(value);
            if (normalized.Contains("cgy") || normalized.Contains("gy"))
            {
                return ReviewUnitCodes.Gray;
            }

            if (normalized.Contains("%"))
            {
                return ReviewUnitCodes.Percent;
            }

            if (normalized.Contains("cm3") ||
                normalized.Contains("cc"))
            {
                return ReviewUnitCodes.CubicCentimeter;
            }

            return ReviewUnitCodes.Text;
        }

        private static ReviewStatusAndSeverity NotEvaluated()
        {
            return new ReviewStatusAndSeverity(
                ReviewStatusCodes.NotEvaluated,
                ReviewSeverityCodes.Info);
        }

        private static bool IsUnsafe(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return false;
            }

            return AbsolutePathPattern.IsMatch(value) ||
                   DicomUidPattern.IsMatch(value.Trim());
        }

        private static string Normalize(string value)
        {
            return (value ?? string.Empty).Trim().ToLowerInvariant();
        }

        private static string ToStableId(string value)
        {
            var result = new StringBuilder();
            bool previousWasSeparator = false;
            foreach (char character in (value ?? string.Empty).Trim())
            {
                if (char.IsLetterOrDigit(character))
                {
                    result.Append(char.ToLowerInvariant(character));
                    previousWasSeparator = false;
                }
                else if (!previousWasSeparator && result.Length > 0)
                {
                    result.Append('-');
                    previousWasSeparator = true;
                }
            }

            return result.ToString().Trim('-');
        }
    }
}
