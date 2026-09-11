using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ClearPlan.Core.Constraints
{
    public static class ConstraintTableSelector
    {
        public static ConstraintTableSelection Select(
            IEnumerable<ConstraintTableDefinition> tables,
            PlanConstraintContext context)
        {
            context = context ?? new PlanConstraintContext();
            var selection = new ConstraintTableSelection();
            var activeTables = (tables ??
                Enumerable.Empty<ConstraintTableDefinition>())
                .Where(item => item != null && item.Active).ToList();
            var labels = new HashSet<string>((context.PrescriptionLabels ?? new List<string>())
                .Where(label => !string.IsNullOrWhiteSpace(label)).Select(label => label.Trim()), StringComparer.OrdinalIgnoreCase);
            var linked = activeTables.Where(table => (table.PrescriptionLabels ?? new List<string>())
                .Any(label => label != null && labels.Contains(label.Trim())) ||
                (!string.IsNullOrEmpty(TableCode(table.DisplayName)) && labels.Any(label => TableCode(label) == TableCode(table.DisplayName)))).ToList();
            foreach (ConstraintTableDefinition table in activeTables)
            {
                ConstraintTableSelectionCandidate candidate = Score(table, context);
                if (candidate != null)
                {
                    if (linked.Contains(table)) { candidate.Score += 200; candidate.Reasons.Add("exact linked prescription label or anchored RefDB table code"); }
                    selection.Candidates.Add(candidate);
                }
            }

            selection.Candidates = selection.Candidates
                .OrderByDescending(candidate => candidate.HasMatchingFractionScope)
                .ThenByDescending(candidate => candidate.StructureHits)
                .ThenByDescending(candidate => candidate.Score)
                .ThenBy(candidate => candidate.Table.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(candidate => candidate.Table.TableId, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (selection.Candidates.Count == 0)
            {
                selection.RequiresConfirmation = true;
                selection.Reason = linked.Count > 0
                    ? "The linked prescription matches RefDB, but its table is incompatible with the active plan fractionation/dose. No full-course goals were applied to a partial plan."
                    : "No compatible constraint table was found.";
                return selection;
            }

            ConstraintTableSelectionCandidate best = selection.Candidates[0];
            if (!context.FractionCount.HasValue &&
                (best.Table.FractionCountMinimum.HasValue || best.Table.FractionCountMaximum.HasValue))
            {
                selection.RequiresConfirmation = true;
                selection.Reason = "Fraction count unavailable: confirm treatment scope before using a fraction-specific table.";
                return selection;
            }
            if (linked.Count > 0 && !selection.Candidates.Any(candidate => linked.Contains(candidate.Table)))
            {
                selection.RequiresConfirmation = true;
                selection.Reason = "The linked prescription table is incompatible with the active plan. Confirm scope before evaluating an alternative.";
                return selection;
            }
            if (linked.Count > 0 && context.FractionCount.HasValue && context.PrescriptionFractionCount.HasValue &&
                context.FractionCount.Value != context.PrescriptionFractionCount.Value)
            {
                selection.RequiresConfirmation = true;
                selection.Reason = "The linked prescription identifies " + best.Table.DisplayName +
                    ", but prescription fractions (" + context.PrescriptionFractionCount + ") differ from active-plan fractions (" + context.FractionCount +
                    "). Confirm the assessment scope; no full-course limits are automatically applied to a partial plan.";
                return selection;
            }
            IList<ConstraintTableSelectionCandidate> tied = selection.Candidates
                .Where(candidate => candidate.HasMatchingFractionScope == best.HasMatchingFractionScope &&
                    candidate.StructureHits == best.StructureHits && candidate.Score == best.Score)
                .ToList();
            int margin = selection.Candidates.Count == 1
                ? int.MaxValue
                : best.Score - selection.Candidates[1].Score;
            bool highConfidence = selection.Candidates.Count == 1 ||
                                  best.StructureHits > selection.Candidates[1].StructureHits ||
                                  margin >= 10 ||
                                  best.Reasons.Any(reason =>
                                      reason.IndexOf("exact fraction", StringComparison.OrdinalIgnoreCase) >= 0);
            if (tied.Count == 1 && highConfidence)
            {
                if (best.Table.RequiresConfirmation)
                {
                    selection.RequiresConfirmation = true;
                    selection.Reason = "The configured table requires explicit confirmation of its clinical scope. " + string.Join("; ", best.Reasons);
                    return selection;
                }
                selection.SelectedTable = best.Table;
                selection.Reason = string.Join("; ", best.Reasons);
                return selection;
            }

            selection.RequiresConfirmation = true;
            selection.Reason = tied.Count > 1
                ? "Multiple constraint tables have the same score."
                : "The best constraint table has low confidence.";
            return selection;
        }

        private static string TableCode(string label)
        {
            var match = Regex.Match(label ?? string.Empty, @"^\s*([A-Za-z][A-Za-z0-9_+\-]{1,31}):");
            return match.Success ? match.Groups[1].Value.ToUpperInvariant() : string.Empty;
        }

        private static ConstraintTableSelectionCandidate Score(
            ConstraintTableDefinition table,
            PlanConstraintContext context)
        {
            if (table.IsPlanSum != context.IsPlanSum)
            {
                return null;
            }

            var candidate = new ConstraintTableSelectionCandidate
            {
                Table = table,
                Score = 100
            };
            candidate.Reasons.Add(context.IsPlanSum
                ? "PlanSum-compatible"
                : "single-plan compatible");

            if (!ScoreIntegerRange(
                context.FractionCount,
                table.FractionCountMinimum,
                table.FractionCountMaximum,
                100,
                70,
                "fraction count",
                candidate))
            {
                return null;
            }
            candidate.HasMatchingFractionScope = context.FractionCount.HasValue &&
                (table.FractionCountMinimum.HasValue || table.FractionCountMaximum.HasValue);

            if (!ScoreDecimalRange(
                context.DosePerFractionGy,
                table.DosePerFractionMinimumGy,
                table.DosePerFractionMaximumGy,
                40,
                "dose per fraction",
                candidate))
            {
                return null;
            }

            if (!ScoreDecimalRange(
                context.TotalDoseGy,
                table.TotalDoseMinimumGy,
                table.TotalDoseMaximumGy,
                30,
                "total dose",
                candidate))
            {
                return null;
            }

            AddStructureCoverage(table, context, candidate);
            AddHintScore(context.SiteHint, table.Site, "site", candidate);
            AddHintScore(context.RegimeHint, table.Regime, "regime", candidate);
            return candidate;
        }

        private static bool ScoreIntegerRange(
            int? actual,
            int? minimum,
            int? maximum,
            int exactScore,
            int rangeScore,
            string label,
            ConstraintTableSelectionCandidate candidate)
        {
            if (!minimum.HasValue && !maximum.HasValue)
            {
                candidate.Score += 10;
                candidate.Reasons.Add("generic " + label);
                return true;
            }

            if (!actual.HasValue)
            {
                candidate.Reasons.Add(label + " unavailable");
                return true;
            }

            int lower = minimum ?? int.MinValue;
            int upper = maximum ?? int.MaxValue;
            if (actual.Value < lower || actual.Value > upper)
            {
                return false;
            }

            if (minimum.HasValue &&
                maximum.HasValue &&
                minimum.Value == maximum.Value &&
                actual.Value == minimum.Value)
            {
                candidate.Score += exactScore;
                candidate.Reasons.Add("exact fraction count");
            }
            else
            {
                candidate.Score += rangeScore;
                candidate.Reasons.Add(label + " within range");
            }

            return true;
        }

        private static bool ScoreDecimalRange(
            decimal? actual,
            decimal? minimum,
            decimal? maximum,
            int score,
            string label,
            ConstraintTableSelectionCandidate candidate)
        {
            if (!minimum.HasValue && !maximum.HasValue)
            {
                return true;
            }

            if (!actual.HasValue)
            {
                candidate.Reasons.Add(label + " unavailable");
                return true;
            }

            decimal lower = minimum ?? decimal.MinValue;
            decimal upper = maximum ?? decimal.MaxValue;
            if (actual.Value < lower || actual.Value > upper)
            {
                return false;
            }

            candidate.Score += score;
            candidate.Reasons.Add(label + " within range");
            return true;
        }

        private static void AddStructureCoverage(
            ConstraintTableDefinition table,
            PlanConstraintContext context,
            ConstraintTableSelectionCandidate candidate)
        {
            var required = (table.Constraints ?? new List<ConstraintDefinition>())
                .Where(constraint => constraint != null)
                .GroupBy(constraint => StructureAliasResolver.NormalizeName(
                    new[] { constraint.StructureName, constraint.RawStructureName, constraint.StructureId }
                        .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))))
                .Where(group => !string.IsNullOrWhiteSpace(group.Key)).Select(group => group.First()).ToList();
            if (required.Count == 0)
            {
                return;
            }

            var available = (context.StructureIds ?? new List<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value)).Distinct(StringComparer.OrdinalIgnoreCase)
                .Select(id => new StructureCandidate { Id = id }).ToList();
            var hits = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var constraint in required)
            {
                var match = StructureAliasResolver.Resolve(constraint, context.StructureDefinitions, available);
                if (!match.IsAmbiguous && !string.IsNullOrWhiteSpace(match.CandidateId)) hits.Add(match.CandidateId);
                // Legacy tables can specify an ID without a display name or catalog definition.
                if (!match.IsAmbiguous && string.IsNullOrWhiteSpace(match.CandidateId) &&
                    string.IsNullOrWhiteSpace(constraint.StructureName) && string.IsNullOrWhiteSpace(constraint.RawStructureName) &&
                    !(context.StructureDefinitions ?? new List<StructureDefinition>()).Any(definition => definition != null &&
                        string.Equals(definition.StructureId, constraint.StructureId, StringComparison.OrdinalIgnoreCase)))
                {
                    var exact = available.Where(item => StructureAliasResolver.NormalizeName(item.Id) ==
                        StructureAliasResolver.NormalizeName(constraint.StructureId)).ToList();
                    if (exact.Count == 1) hits.Add(exact[0].Id);
                }
            }
            int matches = hits.Count;
            candidate.StructureHits = matches;
            candidate.Reasons.Add(string.Format(
                "distinct structure hits {0}/{1}",
                matches,
                required.Count));
        }

        private static void AddHintScore(
            string hint,
            string value,
            string label,
            ConstraintTableSelectionCandidate candidate)
        {
            if (!string.IsNullOrWhiteSpace(hint) &&
                !string.IsNullOrWhiteSpace(value) &&
                string.Equals(hint.Trim(), value.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                candidate.Score += 8;
                candidate.Reasons.Add(label + " hint match");
            }
        }
    }
}
