using System;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewComparisonTests
    {
        public static void PreservesUnitsMissingAndChangedGoals()
        {
            var reference = SyntheticScenarioFactory.Create("baseline-pass");
            var current = SyntheticScenarioFactory.Create("baseline-pass");
            current.PqmRows[0].AchievedValue -= 3;
            current.PqmRows[0].Goal -= 2;
            var rows = ReviewComparison.CompareConstraints(reference, current);
            var changed = rows.Single(r => r.Objective == current.PqmRows[0].Objective && r.Structure == current.PqmRows[0].TemplateStructure);
            TestAssert.Equal(-3.0, changed.Difference.Value);
            TestAssert.True(changed.Note.Contains("Grenzwerte"));
            current.PqmRows[0].Unit = "cm3";
            rows = ReviewComparison.CompareConstraints(reference, current);
            TestAssert.Equal(7, rows.Count);
            TestAssert.Equal(2, rows.Count(r => r.Reference == null || r.Current == null));
            TestAssert.True(rows.Where(r => r.Reference == null || r.Current == null).All(r => r.Difference == null));
        }
        public static void RejectsAmbiguityMappingAndModeMixing()
        {
            var first = SyntheticScenarioFactory.Create("baseline-pass");
            var next = SyntheticScenarioFactory.Create("baseline-pass");
            next.PqmRows.Add(next.PqmRows[0]);
            var rows = ReviewComparison.CompareConstraints(first, next);
            TestAssert.True(rows.Where(r => r.Note.Contains("Mehrdeutig")).All(r => r.Difference == null));
            next.PqmRows.RemoveAt(next.PqmRows.Count - 1);
            next.PqmRows[0].ResolvedStructureId = "AnotherTarget";
            TestAssert.True(ReviewComparison.CompareConstraints(first, next).Single(r => r.Objective == next.PqmRows[0].Objective && r.Structure == next.PqmRows[0].TemplateStructure).Difference == null);
            next.Synthetic = false;
            TestAssert.Throws<ArgumentException>(() => ReviewComparison.CompareConstraints(first, next));
        }
        public static void RetainsSessionReferenceAndClearsAcrossModes()
        {
            var current = SyntheticScenarioFactory.Create("baseline-pass");
            var controller = new ReviewWorkspaceHostController(() => current, null);
            Exception failure;
            TestAssert.True(controller.TryRefresh(out failure));
            controller.CurrentViewModel.Comparison.PinReferenceCommand.Execute(null);
            current = SyntheticScenarioFactory.Create("target-underdose");
            TestAssert.True(controller.TryRefresh(out failure));
            TestAssert.True(controller.CurrentViewModel.Comparison.HasReference);
            TestAssert.Equal("baseline-pass", controller.CurrentViewModel.Comparison.ReferenceSnapshot.ScenarioId);
            TestAssert.True(controller.CurrentViewModel.Comparison.Constraints.Any(r => r.Difference.StartsWith("-")));
            controller.CurrentViewModel.Comparison.ClearReferenceCommand.Execute(null);
            TestAssert.True(controller.TryRefresh(out failure));
            TestAssert.False(controller.CurrentViewModel.Comparison.HasReference);
            controller.Dispose();
        }
        public static void MissingValuesRemainMissing()
        {
            var first = SyntheticScenarioFactory.Create("baseline-pass");
            var next = SyntheticScenarioFactory.Create("baseline-pass");
            next.PqmRows[0].AchievedValue = null;
            TestAssert.True(ReviewComparison.CompareConstraints(first, next).First(r => r.Current == next.PqmRows[0]).Difference == null);
            TestAssert.Equal("—", PlanAnalysisViewModel.Number(double.NaN));
            TestAssert.Equal("—", PlanAnalysisViewModel.Number(double.PositiveInfinity));
        }

        public static void RetainsReferenceAcrossReplacementHost()
        {
            Exception failure;
            var first = new ReviewWorkspaceHostController(() => SyntheticScenarioFactory.Create("baseline-pass"), null);
            TestAssert.True(first.TryRefresh(out failure));
            first.CurrentViewModel.Comparison.PinReferenceCommand.Execute(null);
            var reference = first.CurrentViewModel.Comparison.ReferenceSnapshot;
            var replacement = new ReviewWorkspaceHostController(() => SyntheticScenarioFactory.Create("target-underdose"), null, reference);
            first.Dispose();
            TestAssert.True(replacement.TryRefresh(out failure));
            TestAssert.True(replacement.CurrentViewModel.Comparison.HasReference, "The replacement host must retain the pinned reference.");
            TestAssert.Equal("baseline-pass", replacement.CurrentViewModel.Comparison.ReferenceSnapshot.ScenarioId);
            TestAssert.True(replacement.CurrentViewModel.Comparison.Constraints.Any(row => row.Difference.StartsWith("-")));
            var clinical = SyntheticScenarioFactory.Create("baseline-pass"); clinical.Synthetic = false;
            TestAssert.True(replacement.TryShowSnapshot(clinical, out failure));
            TestAssert.False(replacement.CurrentViewModel.Comparison.HasReference, "A mode switch must clear the carried reference.");
            replacement.Dispose();
        }

        public static void ValidatesExtendedPrivacyAndNumbers()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.PlanAnalysis.TargetStructureId = @"C:\private\clinical-data";
            TestAssert.False(ReviewSnapshotValidator.ValidateForSimulator(snapshot).IsValid);
            snapshot.PlanAnalysis.TargetStructureId = "PTV_60";
            snapshot.PlanAnalysis.Pam = double.NaN;
            TestAssert.False(ReviewSnapshotValidator.ValidateForSimulator(snapshot).IsValid);
            snapshot.PlanAnalysis.Pam = 0.4;
            snapshot.PlanImages = SyntheticPlanImageFactory.Create(snapshot.ActivePlanKey);
            snapshot.PlanImages[0].Synthetic = false;
            TestAssert.False(ReviewSnapshotValidator.ValidateForSimulator(snapshot).IsValid);
        }

        public static void RejectsUncommittedOrChangedPlanForReport()
        {
            var first = SyntheticScenarioFactory.Create("baseline-pass"); first.Synthetic = false;
            var next = SyntheticScenarioFactory.Create("target-underdose"); next.Synthetic = false;
            var guard = new ReviewPlanContextGuard();
            TestAssert.False(guard.Matches(first, "fixture-plan-a"));
            guard.Commit(first, "fixture-plan-a");
            TestAssert.True(guard.Matches(first, "fixture-plan-a"));
            TestAssert.False(guard.Matches(first, "fixture-plan-b"), "Old PQM/DVH cannot be paired with a newly selected plan's CT.");
            TestAssert.False(guard.Matches(next, "fixture-plan-a"), "A different uncommitted snapshot cannot export.");
            // A failed replacement never commits; the previous complete context stays usable.
            try { throw new InvalidOperationException("Synthetic replacement failure"); } catch (InvalidOperationException) { }
            TestAssert.True(guard.Matches(first, "fixture-plan-a"));
            guard.Commit(next, "fixture-plan-b");
            TestAssert.False(guard.Matches(first, "fixture-plan-a"));
            TestAssert.True(guard.Matches(next, "fixture-plan-b"));
            TestAssert.False(ReviewSnapshotJson.Serialize(next).Contains("fixture-plan-b"));
            guard.Clear();
            TestAssert.False(guard.Matches(next, "fixture-plan-b"));
            TestAssert.Throws<ArgumentException>(() => guard.Commit(SyntheticScenarioFactory.Create("baseline-pass"), "fixture-plan-a"));
            var sumObject = new object();
            guard.Commit(first, "", sumObject);
            TestAssert.True(guard.Matches(first, "", sumObject), "PlanSum review/export must not require a nonexistent SOP UID.");
            TestAssert.False(guard.Matches(first, "", new object()));
            TestAssert.False(guard.Matches(first, "fixture-plan-a", sumObject));
            guard.Clear(); // Enter demo, then return to the same clinical PlanSum.
            guard.Commit(next, "", sumObject);
            TestAssert.True(guard.Matches(next, "", sumObject));
        }
    }
}
