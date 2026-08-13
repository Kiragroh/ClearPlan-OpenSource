using System;
using System.Collections.Generic;
using System.Linq;

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
            foreach (ConstraintTableDefinition table in (tables ??
                Enumerable.Empty<ConstraintTableDefinition>())
                .Where(item => item != null && item.Active))
            {
                ConstraintTableSelectionCandidate candidate = Score(table, context);
                if (candidate != null)
                {
                    selection.Candidates.Add(candidate);
                }
            }

            selection.Candidates = selection.Candidates
                .OrderByDescending(candidate => candidate.Score)
                .ThenBy(candidate => candidate.Table.DisplayName, StringComparer.OrdinalIgnoreCase)
                .ThenBy(candidate => candidate.Table.TableId, StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (selection.Candidates.Count == 0)
            {
                selection.Reason = "No compatible constraint table was found.";
                return selection;
            }

            ConstraintTableSelectionCandidate best = selection.Candidates[0];
            IList<ConstraintTableSelectionCandidate> tied = selection.Candidates
                .Where(candidate => candidate.Score == best.Score)
                .ToList();
            int margin = selection.Candidates.Count == 1
                ? int.MaxValue
                : best.Score - selection.Candidates[1].Score;
            bool highConfidence = selection.Candidates.Count == 1 ||
                                  margin >= 10 ||
                                  best.Reasons.Any(reason =>
                                      reason.IndexOf("exact fraction", StringComparison.OrdinalIgnoreCase) >= 0);
            if (tied.Count == 1 && highConfidence)
            {
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
            var required = new HashSet<string>(
                table.Constraints
                    .Select(constraint => constraint.StructureId)
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(StructureAliasResolver.NormalizeName),
                StringComparer.Ordinal);
            if (required.Count == 0)
            {
                return;
            }

            var available = new HashSet<string>(
                (context.StructureIds ?? new List<string>())
                    .Where(value => !string.IsNullOrWhiteSpace(value))
                    .Select(StructureAliasResolver.NormalizeName),
                StringComparer.Ordinal);
            int matches = required.Count(available.Contains);
            int coverageScore = (int)Math.Round(20.0 * matches / required.Count);
            candidate.Score += coverageScore;
            candidate.Reasons.Add(string.Format(
                "structure coverage {0}/{1}",
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
