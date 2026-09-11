using System;
using System.Collections.Generic;
using System.Linq;

namespace ClearPlan.Core.Review
{
    /// <summary>Descriptive, unit-preserving comparison. Never interprets a delta as clinical superiority.</summary>
    public static class ReviewComparison
    {
        public static List<ReviewComparisonRow> CompareConstraints(ReviewSnapshot reference, ReviewSnapshot current)
        {
            if (reference == null || current == null) return new List<ReviewComparisonRow>();
            if (reference.Synthetic != current.Synthetic)
                throw new ArgumentException("Synthetic and clinical snapshots cannot be compared.");
            var left = reference.PqmRows.GroupBy(Key).ToDictionary(g => g.Key, g => g.ToList());
            var right = current.PqmRows.GroupBy(Key).ToDictionary(g => g.Key, g => g.ToList());
            var result = new List<ReviewComparisonRow>();
            foreach (string key in left.Keys.Union(right.Keys).OrderBy(k => k, StringComparer.Ordinal))
            {
                List<ReviewPqmRow> a, b;
                left.TryGetValue(key, out a);
                right.TryGetValue(key, out b);
                int count = Math.Max(a == null ? 0 : a.Count, b == null ? 0 : b.Count);
                bool ambiguous = (a != null && a.Count > 1) || (b != null && b.Count > 1);
                for (int index = 0; index < count; index++)
                {
                    ReviewPqmRow first = a != null && index < a.Count ? a[index] : null;
                    ReviewPqmRow second = b != null && index < b.Count ? b[index] : null;
                    var exemplar = first ?? second;
                    bool comparable = !ambiguous && first != null && second != null &&
                        !string.IsNullOrWhiteSpace(first.ResolvedStructureId) &&
                        !string.IsNullOrWhiteSpace(second.ResolvedStructureId) &&
                        string.Equals(first.ResolvedStructureId, second.ResolvedStructureId, StringComparison.OrdinalIgnoreCase);
                    result.Add(new ReviewComparisonRow
                    {
                        Structure = exemplar.TemplateStructure,
                        Objective = exemplar.Objective,
                        Unit = exemplar.Unit,
                        Reference = first,
                        Current = second,
                        Difference = comparable && Finite(first.AchievedValue) && Finite(second.AchievedValue)
                            ? second.AchievedValue - first.AchievedValue : null,
                        Note = ambiguous ? "Mehrdeutiges Kriterium – keine automatische Differenz" :
                            first == null ? "Nur im aktuellen Plan" : second == null ? "Nur in der Referenz" :
                            !comparable ? "Strukturzuordnung verschieden oder fehlt" :
                            first.Goal != second.Goal || first.Variation != second.Variation || first.Comparator != second.Comparator
                                ? "Grenzwerte verschieden; Werte rein deskriptiv" : "Gleiches Kriterium"
                    });
                }
            }
            return result;
        }

        private static string Key(ReviewPqmRow row)
        {
            // Unit and objective are part of identity: Gy, %, D95 and V95 are not interchangeable.
            return (row.TemplateStructure ?? "").Trim().ToUpperInvariant() + "\u001f" +
                (row.Objective ?? "").Trim().ToUpperInvariant() + "\u001f" + (row.Unit ?? "").Trim();
        }

        public static bool Finite(double? value)
        {
            return value.HasValue && !double.IsNaN(value.Value) && !double.IsInfinity(value.Value);
        }
    }

    public sealed class ReviewComparisonRow
    {
        public string Structure { get; set; }
        public string Objective { get; set; }
        public string Unit { get; set; }
        public ReviewPqmRow Reference { get; set; }
        public ReviewPqmRow Current { get; set; }
        public double? Difference { get; set; }
        public string Note { get; set; }
    }
}
