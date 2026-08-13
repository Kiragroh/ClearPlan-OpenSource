using System;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewWorkspaceHostControllerTests
    {
        public static void DispatchesVisibleActionsExactlyOnce()
        {
            int reportCalls = 0;
            int resetCalls = 0;
            int mappingCalls = 0;
            object mappingParameter = null;
            var bindings = new ReviewWorkspaceActionBindings
            {
                Report = (sender, args) => reportCalls++,
                ResetDvh = (sender, args) => resetCalls++,
                ApplyStructureMapping = (sender, args) =>
                {
                    mappingCalls++;
                    mappingParameter = args.Parameter;
                }
            };
            using (var controller = new ReviewWorkspaceHostController(
                () => SyntheticScenarioFactory.Create("baseline-pass"),
                bindings))
            {
                Exception failure;
                TestAssert.True(controller.TryRefresh(out failure));
                TestAssert.True(failure == null);

                controller.CurrentViewModel.ReportCommand.Execute(null);
                controller.CurrentViewModel.ResetDvhCommand.Execute(null);
                controller.CurrentViewModel.ApplyStructureMappingCommand.Execute(
                    "mapping-parameter");

                TestAssert.Equal(1, reportCalls);
                TestAssert.Equal(1, resetCalls);
                TestAssert.Equal(1, mappingCalls);
                TestAssert.Equal("mapping-parameter", mappingParameter);
            }
        }

        public static void KeepsLastGoodViewModelWhenRefreshFails()
        {
            bool fail = false;
            using (var controller = new ReviewWorkspaceHostController(
                () =>
                {
                    if (fail)
                    {
                        throw new InvalidOperationException("synthetic failure");
                    }

                    return SyntheticScenarioFactory.Create("baseline-pass");
                },
                new ReviewWorkspaceActionBindings()))
            {
                Exception firstFailure;
                TestAssert.True(controller.TryRefresh(out firstFailure));
                ReviewWorkspaceViewModel first = controller.CurrentViewModel;

                fail = true;
                Exception secondFailure;
                TestAssert.False(controller.TryRefresh(out secondFailure));
                TestAssert.True(secondFailure is InvalidOperationException);
                TestAssert.True(ReferenceEquals(
                    first,
                    controller.CurrentViewModel));
            }
        }

        public static void UnsubscribesActionsWhenDisposed()
        {
            int reportCalls = 0;
            var bindings = new ReviewWorkspaceActionBindings
            {
                Report = (sender, args) => reportCalls++
            };
            var controller = new ReviewWorkspaceHostController(
                () => SyntheticScenarioFactory.Create("baseline-pass"),
                bindings);
            Exception failure;
            TestAssert.True(controller.TryRefresh(out failure));
            ReviewWorkspaceViewModel viewModel =
                controller.CurrentViewModel;

            controller.Dispose();
            viewModel.ReportCommand.Execute(null);

            TestAssert.Equal(0, reportCalls);
        }

        public static void UsesSingularHintLabelForOneVariation()
        {
            var viewModel = new ReviewWorkspaceViewModel(
                SyntheticScenarioFactory.Create("target-underdose"));

            TestAssert.Equal(
                "4 erfüllt · 1 Hinweis · 1 nicht erfüllt",
                viewModel.PqmSummary);
        }
    }
}
