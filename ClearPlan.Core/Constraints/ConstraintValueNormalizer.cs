using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Constraints
{
    public static class ConstraintValueNormalizer
    {
        public static string NormalizeComparator(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized == "≤")
            {
                return "<=";
            }

            if (normalized == "≥")
            {
                return ">=";
            }

            return normalized;
        }

        public static string NormalizeUnit(string value)
        {
            string normalized = (value ?? string.Empty).Trim();
            if (normalized.Equals("cm³", StringComparison.OrdinalIgnoreCase) ||
                normalized.Equals("cc", StringComparison.OrdinalIgnoreCase))
            {
                return "cc";
            }

            if (normalized.Equals("gy", StringComparison.OrdinalIgnoreCase))
            {
                return "Gy";
            }

            if (normalized.StartsWith("cm (", StringComparison.OrdinalIgnoreCase))
            {
                return "cm";
            }

            return normalized;
        }

        public static decimal? ParseNullableDecimal(object rawValue)
        {
            if (rawValue == null)
            {
                return null;
            }

            string value = Convert.ToString(rawValue, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(value))
            {
                return null;
            }

            value = value.Trim();
            if (value.IndexOf(',') >= 0 && value.IndexOf('.') < 0)
            {
                value = value.Replace(',', '.');
            }

            decimal parsed;
            if (!decimal.TryParse(
                value,
                NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                CultureInfo.InvariantCulture,
                out parsed))
            {
                throw new FormatException(string.Format(
                    CultureInfo.InvariantCulture,
                    "Constraint value '{0}' is not a valid invariant decimal.",
                    value));
            }

            return parsed;
        }

        public static int? ParseNullableInteger(object rawValue)
        {
            decimal? parsed = ParseNullableDecimal(rawValue);
            if (!parsed.HasValue)
            {
                return null;
            }

            if (decimal.Truncate(parsed.Value) != parsed.Value)
            {
                throw new FormatException("Constraint value is not a whole number.");
            }

            return decimal.ToInt32(parsed.Value);
        }

        public static string NormalizeMetric(string value)
        {
            string normalized = Regex.Replace((value ?? string.Empty).Trim(), @"\s+", string.Empty);
            normalized = normalized.Replace(',', '.').Replace("cm³", "cc");
            normalized = Regex.Replace(normalized, "gy", "Gy", RegexOptions.IgnoreCase);
            normalized = Regex.Replace(normalized, "cc", "cc", RegexOptions.IgnoreCase);

            if (normalized.Equals("dmean", StringComparison.OrdinalIgnoreCase))
            {
                return "Dmean";
            }

            if (normalized.Equals("mean", StringComparison.OrdinalIgnoreCase))
            {
                return "Dmean";
            }

            if (normalized.Equals("max", StringComparison.OrdinalIgnoreCase))
            {
                return "Dmax";
            }

            if (normalized.Equals("min", StringComparison.OrdinalIgnoreCase))
            {
                return "Dmin";
            }

            if (normalized.StartsWith("cv", StringComparison.OrdinalIgnoreCase))
            {
                return "CV" + normalized.Substring(2);
            }

            if (normalized.StartsWith("d", StringComparison.OrdinalIgnoreCase))
            {
                return "D" + normalized.Substring(1);
            }

            if (normalized.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                return "V" + normalized.Substring(1);
            }

            return normalized;
        }
    }
}
