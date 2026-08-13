using System;
using System.Collections.Generic;
using System.Linq;
using ClearPlan.Core.Review;

namespace ClearPlan.Simulator
{
    public static class SimulatorSnapshotActions
    {
        public static string GetOpenPlanStatus(
            ReviewSnapshot snapshot,
            string planKey)
        {
            RequireSyntheticSnapshot(snapshot);
            ReviewPlanRow plan = (snapshot.Plans ??
                                  new List<ReviewPlanRow>())
                .SingleOrDefault(
                    row => string.Equals(
                        row.PlanKey,
                        planKey,
                        StringComparison.Ordinal));
            if (plan == null)
            {
                throw new ArgumentException(
                    "The synthetic plan is not part of the current scenario.",
                    "planKey");
            }

            if (string.Equals(
                plan.PlanKey,
                snapshot.ActivePlanKey,
                StringComparison.Ordinal))
            {
                return "Synthetisch aktiver Plan · " +
                       plan.DisplayLabel;
            }

            ReviewPlanRow activePlan = snapshot.Plans.Single(
                row => string.Equals(
                    row.PlanKey,
                    snapshot.ActivePlanKey,
                    StringComparison.Ordinal));
            return "Synthetischer Vergleichsplan · " +
                   plan.DisplayLabel +
                   " · aktiv bleibt " +
                   activePlan.DisplayLabel;
        }

        public static string ApplyStructureMapping(
            ReviewSnapshot snapshot,
            string stableId,
            string selectedStructureId)
        {
            RequireSyntheticSnapshot(snapshot);
            ReviewStructureMapping mapping =
                (snapshot.StructureMappings ??
                 new List<ReviewStructureMapping>())
                    .SingleOrDefault(
                        row => string.Equals(
                            row.StableId,
                            stableId,
                            StringComparison.Ordinal));
            if (mapping == null)
            {
                throw new ArgumentException(
                    "The structure mapping is not part of the current scenario.",
                    "stableId");
            }
            if (string.IsNullOrWhiteSpace(selectedStructureId) ||
                !(mapping.AvailableStructureIds ??
                  new List<string>())
                    .Contains(
                        selectedStructureId,
                        StringComparer.Ordinal))
            {
                throw new ArgumentException(
                    "Choose one of the synthetic structures offered by the scenario.",
                    "selectedStructureId");
            }

            mapping.SelectedStructureId = selectedStructureId;
            mapping.Status = ReviewStatusCodes.Pass;
            mapping.Message =
                "Synthetische Zuordnung übernommen: " +
                mapping.TemplateStructure +
                " → " +
                selectedStructureId;
            return "Zuordnung übernommen · " +
                   mapping.TemplateStructure +
                   " → " +
                   selectedStructureId;
        }

        public static int SynchronizeDvhSelections(
            ReviewSnapshot snapshot,
            IEnumerable<KeyValuePair<string, bool>> selections)
        {
            RequireSyntheticSnapshot(snapshot);
            if (selections == null)
            {
                throw new ArgumentNullException("selections");
            }

            IDictionary<string, bool> byStableId =
                selections.ToDictionary(
                    item => item.Key,
                    item => item.Value,
                    StringComparer.Ordinal);
            foreach (ReviewDvhSeries series in
                     snapshot.DvhSeries ?? new List<ReviewDvhSeries>())
            {
                bool selected;
                if (byStableId.TryGetValue(
                    series.StableId,
                    out selected))
                {
                    series.Selected = selected;
                }
            }

            return (snapshot.DvhSeries ??
                    new List<ReviewDvhSeries>())
                .Count(series => series.Selected);
        }

        private static void RequireSyntheticSnapshot(
            ReviewSnapshot snapshot)
        {
            if (snapshot == null)
            {
                throw new ArgumentNullException("snapshot");
            }
            if (!snapshot.Synthetic)
            {
                throw new InvalidOperationException(
                    "Simulator actions require a synthetic snapshot.");
            }
        }
    }
}
