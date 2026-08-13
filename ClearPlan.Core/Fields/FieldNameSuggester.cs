using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace ClearPlan.Core.Fields
{
    public static class FieldNameSuggester
    {
        private const double AngleToleranceDegrees = 0.5;

        public static IList<FieldNameSuggestion> Suggest(
            IEnumerable<BeamNamingInput> beams)
        {
            return SuggestInternal(null, beams, false);
        }

        public static IList<FieldNameSuggestion> Suggest(
            string planId,
            IEnumerable<BeamNamingInput> beams)
        {
            return SuggestInternal(planId, beams, true);
        }

        private static IList<FieldNameSuggestion> SuggestInternal(
            string planId,
            IEnumerable<BeamNamingInput> beams,
            bool evaluateIds)
        {
            IList<BeamNamingInput> ordered = (beams ??
                Enumerable.Empty<BeamNamingInput>())
                .Where(beam => beam != null)
                .OrderBy(beam => beam.TreatmentOrderIndex.HasValue ? 0 : 1)
                .ThenBy(beam => beam.TreatmentOrderIndex ?? int.MaxValue)
                .ThenBy(beam => beam.BeamNumber)
                .ThenBy(beam => beam.CurrentId, StringComparer.OrdinalIgnoreCase)
                .ToList();
            var baseNames = ordered.ToDictionary(
                beam => beam,
                BuildBaseName);
            var counts = baseNames.Values
                .GroupBy(name => name, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(),
                    StringComparer.OrdinalIgnoreCase);
            var indexes = new Dictionary<string, int>(
                StringComparer.OrdinalIgnoreCase);
            string idPrefix = evaluateIds
                ? GetTreatmentIdPrefix(planId, ordered)
                : string.Empty;
            var suggestions = new List<FieldNameSuggestion>();
            for (int displayOrder = 0; displayOrder < ordered.Count; displayOrder++)
            {
                BeamNamingInput beam = ordered[displayOrder];
                string baseName = baseNames[beam];
                string suggested = baseName;
                if (counts[baseName] > 1)
                {
                    int duplicateIndex;
                    if (!indexes.TryGetValue(baseName, out duplicateIndex))
                    {
                        duplicateIndex = 0;
                    }

                    suggested = AddDuplicateSuffix(baseName, duplicateIndex);
                    indexes[baseName] = duplicateIndex + 1;
                }

                string expectedId = evaluateIds
                    ? TruncateId(idPrefix +
                                 (displayOrder + 1).ToString(
                                     "00",
                                     CultureInfo.InvariantCulture))
                    : beam.CurrentId ?? string.Empty;
                bool idWouldChange = evaluateIds &&
                                     !string.Equals(
                                         beam.CurrentId ?? string.Empty,
                                         expectedId,
                                         StringComparison.OrdinalIgnoreCase);
                bool nameWouldChange = !string.Equals(
                    beam.CurrentName ?? string.Empty,
                    suggested,
                    StringComparison.Ordinal);
                suggestions.Add(new FieldNameSuggestion
                {
                    CurrentId = beam.CurrentId ?? string.Empty,
                    ExpectedId = expectedId,
                    CurrentName = beam.CurrentName ?? string.Empty,
                    BeamNumber = beam.BeamNumber,
                    DisplayOrder = displayOrder + 1,
                    SuggestedName = suggested,
                    IdWouldChange = idWouldChange,
                    NameWouldChange = nameWouldChange,
                    WouldChange = idWouldChange || nameWouldChange
                });
            }

            return suggestions;
        }

        private static string GetTreatmentIdPrefix(
            string planId,
            IList<BeamNamingInput> ordered)
        {
            BeamNamingInput first = ordered.FirstOrDefault();
            string currentId = first == null
                ? string.Empty
                : (first.CurrentId ?? string.Empty).Trim();
            if (!string.IsNullOrWhiteSpace(currentId) &&
                !currentId.StartsWith("TMP", StringComparison.OrdinalIgnoreCase) &&
                currentId.IndexOf(' ') < 0)
            {
                int end = currentId.Length;
                while (end > 0 && char.IsDigit(currentId[end - 1]))
                {
                    end--;
                }

                string currentPrefix = end > 0
                    ? currentId.Substring(0, end)
                    : currentId;
                if (!string.IsNullOrWhiteSpace(currentPrefix))
                {
                    return TruncatePrefix(currentPrefix);
                }
            }

            string source = planId ?? string.Empty;
            int prefixLength = 0;
            while (prefixLength < source.Length &&
                   char.IsLetterOrDigit(source[prefixLength]))
            {
                prefixLength++;
            }

            string planPrefix = prefixLength > 0
                ? source.Substring(0, prefixLength)
                : "SET";
            return TruncatePrefix(planPrefix);
        }

        private static string TruncatePrefix(string value)
        {
            string prefix = string.IsNullOrWhiteSpace(value)
                ? "SET"
                : value.Trim();
            return prefix.Length <= 12
                ? prefix
                : prefix.Substring(0, 12);
        }

        private static string TruncateId(string value)
        {
            return value.Length <= 16
                ? value
                : value.Substring(0, 16);
        }

        private static string BuildBaseName(BeamNamingInput beam)
        {
            bool isArc = !SameAngle(
                beam.GantryStartAngle,
                beam.GantryStopAngle);
            string anglePart = isArc
                ? AngleText(beam.GantryStartAngle) + "-" +
                  AngleText(beam.GantryStopAngle)
                : AngleText(beam.GantryStartAngle);
            string tablePart = SameAngle(beam.PatientSupportAngle, 0)
                ? string.Empty
                : " T" + AngleText(beam.PatientSupportAngle);
            string directionPart = isArc
                ? " " + ArcDirectionToken(beam.GantryDirection)
                : string.Empty;

            return anglePart + tablePart + directionPart;
        }

        private static string AddDuplicateSuffix(string baseName, int index)
        {
            string letter = DuplicateLetter(index);
            if (baseName.EndsWith(" UZ", StringComparison.OrdinalIgnoreCase) ||
                baseName.EndsWith(" GUZ", StringComparison.OrdinalIgnoreCase))
            {
                return baseName + letter;
            }

            return baseName + " UZ" + letter;
        }

        private static string DuplicateLetter(int index)
        {
            if (index >= 0 && index < 26)
            {
                return ((char)('a' + index)).ToString();
            }

            return (index + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static string ArcDirectionToken(string gantryDirection)
        {
            string direction = (gantryDirection ?? string.Empty)
                .ToUpperInvariant();
            if (direction.Contains("COUNTER") || direction.Contains("CCW"))
            {
                return "GUZ";
            }

            if (direction.Contains("CLOCKWISE") || direction.Contains("CW"))
            {
                return "UZ";
            }

            return "UZ";
        }

        private static bool SameAngle(double first, double second)
        {
            return Math.Abs(
                NormalizeAngle(first) -
                NormalizeAngle(second)) <= AngleToleranceDegrees;
        }

        private static string AngleText(double angle)
        {
            return Math.Round(NormalizeAngle(angle), 0)
                .ToString("0", CultureInfo.InvariantCulture);
        }

        private static double NormalizeAngle(double angle)
        {
            double normalized = angle % 360.0;
            return normalized < 0 ? normalized + 360.0 : normalized;
        }
    }
}
