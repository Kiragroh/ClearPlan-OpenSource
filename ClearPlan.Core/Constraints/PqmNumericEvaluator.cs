using System;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Constraints
{
    /// <summary>Strict PQM decimals: comma or dot are decimal separators, never grouping.</summary>
    public static class PqmNumericEvaluator
    {
        private const string Number = @"[-+]?(?:\d+(?:[.,]\d+)?|[.,]\d+)";
        private static readonly Regex AchievedPattern = new Regex(
            @"^\s*(?<value>" + Number + @")\s*(?:Gy|cGy|%|cc|cm3|cm³)?\s*$",
            RegexOptions.CultureInvariant);
        private static readonly Regex GoalPattern = new Regex(
            @"^\s*(?<operator><=|>=|<|>|=)\s*(?<value>" + Number + @")\s*$",
            RegexOptions.CultureInvariant);

        public static bool TryParseNumber(string text, out double value)
        {
            value = 0;
            try
            {
                decimal? parsed = ConstraintValueNormalizer.ParseNullableDecimal(text);
                if (!parsed.HasValue) return false;
                value = (double)parsed.Value;
                return true;
            }
            catch (FormatException) { return false; }
        }

        public static bool TryParseAchieved(string text, out double value)
        {
            value = 0;
            Match match = AchievedPattern.Match(text ?? string.Empty);
            return match.Success && TryParseNumber(match.Groups["value"].Value, out value) && value >= 0;
        }

        public static bool TryParseGoal(string text, out string comparator, out double value)
        {
            comparator = string.Empty;
            value = 0;
            Match match = GoalPattern.Match(text ?? string.Empty);
            if (!match.Success || !TryParseNumber(match.Groups["value"].Value, out value) || value < 0) return false;
            comparator = match.Groups["operator"].Value;
            return true;
        }

        public static string Evaluate(string achieved, string goal, string variation)
        {
            double actual, limit, allowed = 0;
            string comparator;
            bool hasVariation = !string.IsNullOrWhiteSpace(variation);
            if (!TryParseAchieved(achieved, out actual) || !TryParseGoal(goal, out comparator, out limit) ||
                (hasVariation && (!TryParseNumber(variation, out allowed) || allowed < 0))) return "Not evaluated";
            if (Matches(actual, limit, comparator)) return "Goal";
            return hasVariation && Matches(actual, allowed, comparator) ? "Variation" : "Not met";
        }

        public static string EvaluateForConfirmedScope(string achieved, string goal, string variation, bool scopeConfirmed)
        {
            return scopeConfirmed ? Evaluate(achieved, goal, variation) : "Not evaluated";
        }

        public static bool TryGetGoalRatio(string achieved, string goal, out double ratio)
        {
            ratio = 0;
            double actual, limit;
            string comparator;
            if (!TryParseAchieved(achieved, out actual) || !TryParseGoal(goal, out comparator, out limit) || limit == 0)
                return false;
            ratio = actual / limit * 100;
            return !double.IsNaN(ratio) && !double.IsInfinity(ratio);
        }

        public static bool TryConvertDose(double value, string sourceUnit, string targetUnit, out double converted)
        {
            converted = 0;
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) return false;
            if (sourceUnit == "%" && targetUnit == "%") { converted = value; return true; }
            if ((sourceUnit != "Gy" && sourceUnit != "cGy") || (targetUnit != "Gy" && targetUnit != "cGy"))
                return false;
            converted = value * (sourceUnit == "cGy" ? 0.01 : 1) * (targetUnit == "cGy" ? 100 : 1);
            return !double.IsNaN(converted) && !double.IsInfinity(converted);
        }

        public static bool TryComplementVolume(double volumeAtDose, double organVolumeCc, string volumeUnit,
            double? minimumVolume, out double complement)
        {
            complement = 0;
            if (double.IsNaN(organVolumeCc) || double.IsInfinity(organVolumeCc) || organVolumeCc <= 0 ||
                (volumeUnit != "%" && volumeUnit != "cc")) return false;
            double total = volumeUnit == "%" ? 100 : organVolumeCc;
            if (double.IsNaN(volumeAtDose) || double.IsInfinity(volumeAtDose) || volumeAtDose < 0 || volumeAtDose > total)
                return false;
            if (minimumVolume.HasValue && (double.IsNaN(minimumVolume.Value) || double.IsInfinity(minimumVolume.Value) ||
                minimumVolume.Value < 0 || minimumVolume.Value > total)) return false;
            complement = total - volumeAtDose;
            return true;
        }

        public static bool TryGetComplementGoalRatio(string achieved, string goal, double organVolumeCc,
            string volumeUnit, out double ratio)
        {
            ratio = 0;
            double actual, limit, actualComplement, limitComplement;
            string comparator;
            if (!TryParseAchieved(achieved, out actual) || !TryParseGoal(goal, out comparator, out limit) ||
                !TryComplementVolume(actual, organVolumeCc, volumeUnit, null, out actualComplement) ||
                !TryComplementVolume(limit, organVolumeCc, volumeUnit, null, out limitComplement) || limitComplement == 0)
                return false;
            ratio = actualComplement / limitComplement * 100;
            return !double.IsNaN(ratio) && !double.IsInfinity(ratio);
        }

        private static bool Matches(double actual, double limit, string comparator)
        {
            switch (comparator)
            {
                case "<": return actual < limit;
                case "<=": return actual <= limit;
                case ">": return actual > limit;
                case ">=": return actual >= limit;
                case "=": return actual == limit;
                default: return false;
            }
        }
    }
}
