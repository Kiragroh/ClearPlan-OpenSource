using System;
using System.IO;
using System.Linq;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class NativeWarningsTests
    {
        public static void SeparateSourceAndNoInferredApproval()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.PlanCheckRows.Add(new ReviewCheckRow { CheckCode = "synthetic-native-warning", Category = "Eclipse warnings",
                Status = ReviewStatusCodes.Variation, Severity = ReviewSeverityCodes.Warning, Unit = ReviewUnitCodes.Text,
                Message = "Synthetic native warning fixture" });
            var view = new ReviewWorkspaceViewModel(snapshot);
            var property = view.GetType().GetProperty("WarningRows");
            TestAssert.NotNull(property, "Warnings need their own review-tab collection.");
            TestAssert.Equal(1, ((System.Collections.IEnumerable)property.GetValue(view, null)).Cast<object>().Count());
            var source = File.ReadAllText(Path.Combine("ClearPlan.Script", "Review", "EsapiWarningsBuilder.cs"));
            TestAssert.True(source.Contains("IsValidForPlanApproval(out") && source.Contains("MessageForUser"));
            TestAssert.False(source.Contains("BeginModifications") || source.Contains("SaveModifications"));
            TestAssert.True(source.Contains("approval extensions"));
        }
    }
}
