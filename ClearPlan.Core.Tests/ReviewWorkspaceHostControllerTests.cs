using System;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewWorkspaceHostControllerTests
    {
        public static void AriaActionIsPatientBoundAndDisabledInSandbox()
        {
            var clinical=SyntheticScenarioFactory.Create("baseline-pass"); clinical.Synthetic=false;
            int calls=0;
            using(var controller=new ReviewWorkspaceHostController(()=>clinical,new ReviewWorkspaceActionBindings { AriaUpload=(sender,args)=>calls++ }))
            {
                Exception error; TestAssert.True(controller.TryRefresh(out error));
                var view=controller.CurrentViewModel;
                TestAssert.False(view.AriaUploadCommand.CanExecute(null)); view.AriaUploadCommand.Execute(null); TestAssert.Equal(0,calls);
                view.SetAriaUploadAvailability(true,false,"Ready"); view.AriaUploadCommand.Execute(null); TestAssert.Equal(1,calls);
                view.SetAriaUploadAvailability(true,true,"Preparing"); view.AriaUploadCommand.Execute(null); TestAssert.Equal(1,calls);
                TestAssert.True(controller.TryShowSnapshot(SyntheticScenarioFactory.Create("baseline-pass"),out error));
                controller.CurrentViewModel.SetAriaUploadAvailability(true,false,"Ready");
                TestAssert.False(controller.CurrentViewModel.AriaUploadCommand.CanExecute(null));
                controller.CurrentViewModel.AriaUploadCommand.Execute(null); TestAssert.Equal(1,calls);
                view.SetAriaUploadAvailability(true,false,"Ready"); view.AriaUploadCommand.Execute(null); TestAssert.Equal(1,calls);
            }
        }
        public static void ClinicalHostDoesNotInheritLegacyPlotModelsBeforeLoading()
        {
            string source = System.IO.File.ReadAllText(System.IO.Path.Combine("ClearPlan.Script", "MainView.xaml"));
            var element = System.Text.RegularExpressions.Regex.Match(source, @"<review:ReviewWorkspaceView\b[^>]*>").Value;
            TestAssert.True(element.Contains("DataContext=\"{x:Null}\""),
                "The shared view must not inherit MainViewModel.OverviewPlotModel before Loaded: the legacy plot already owns that OxyPlot model.");
        }

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

        public static void SwitchesBetweenClinicalAndSyntheticSnapshots()
        {
            ReviewSnapshot clinical =
                SyntheticScenarioFactory.Create("baseline-pass");
            clinical.Synthetic = false;
            clinical.ScenarioTitle = "Clinical plan review";
            clinical.ScenarioDescription = "Read-only clinical snapshot";

            using (var controller = new ReviewWorkspaceHostController(
                () => clinical,
                new ReviewWorkspaceActionBindings()))
            {
                Exception failure;
                TestAssert.True(controller.TryRefresh(out failure));
                TestAssert.False(controller.CurrentViewModel.IsSynthetic);

                TestAssert.True(controller.TryShowSnapshot(
                    SyntheticScenarioFactory.Create("mixed-review"),
                    out failure));
                TestAssert.True(controller.CurrentViewModel.IsSynthetic);

                TestAssert.True(controller.TryRefresh(out failure));
                TestAssert.False(controller.CurrentViewModel.IsSynthetic);
            }
        }
    }
}
