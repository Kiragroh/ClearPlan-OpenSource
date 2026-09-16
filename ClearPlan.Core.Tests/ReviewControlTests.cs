using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Input;
using ClearPlan.Core.Review;
using ClearPlan.Core.Simulation;
using ClearPlan.Presentation.ViewModels;

namespace ClearPlan.Core.Tests
{
    internal static class ReviewControlTests
    {
        public static void ReferenceSelectionIsSeparateFromOpening()
        {
            var property = typeof(ReviewWorkspaceViewModel).GetProperty("SelectComparisonPlanCommand");
            var binding = typeof(ReviewWorkspaceActionBindings).GetProperty("SelectComparisonPlan");
            TestAssert.NotNull(property, "Comparison needs its own reference action.");
            TestAssert.NotNull(binding);
            int references = 0, opens = 0;
            var bindings = new ReviewWorkspaceActionBindings { OpenPlan = (s, e) => opens++ };
            binding.SetValue(bindings, new EventHandler<ReviewWorkspaceActionEventArgs>((s, e) => references++));
            using (var host = new ReviewWorkspaceHostController(() => SyntheticScenarioFactory.Create("baseline-pass"), bindings))
            {
                Exception error;
                TestAssert.True(host.TryRefresh(out error));
                var previous = host.CurrentViewModel;
                ((ICommand)property.GetValue(previous)).Execute(previous.Plans.Last());
                TestAssert.Equal(1, references); TestAssert.Equal(0, opens);
                TestAssert.True(host.TryRefresh(out error));
                ((ICommand)property.GetValue(previous)).Execute(previous.Plans.Last());
                TestAssert.Equal(1, references, "Obsolete workspace must be unsubscribed.");
            }
        }

        public static void ExplicitTargetVisibilityAndRefreshAreShared()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.DvhSeries[0].RequiredForTargetReview = true;
            snapshot.DvhSeries[0].Selected = true;
            snapshot.DvhSeries[1].RequiredForTargetReview = false;
            snapshot.DvhSeries[1].Selected = false;
            using (var host = new ReviewWorkspaceHostController(() => snapshot, new ReviewWorkspaceActionBindings()))
            {
                Exception error; TestAssert.True(host.TryRefresh(out error));
                var row = host.CurrentViewModel.DvhSeries[0];
                TestAssert.False(host.CurrentViewModel.PlanImages.IsStructureVisible(snapshot.DvhSeries[1].StructureId),
                    "Initial DVH exclusions must also initialize CT visibility and report options.");
                row.IsSelected = false;
                TestAssert.False(row.IsSelected, "An explicit legend click may hide a target without deleting its DVH or metrics.");
                var method = host.CurrentViewModel.PlanImages.GetType().GetMethod("IsStructureVisible");
                TestAssert.NotNull(method, "CT and DVH need a shared visibility contract.");
                TestAssert.False((bool)method.Invoke(host.CurrentViewModel.PlanImages, new object[] { row.StructureId }));
                TestAssert.True(host.TryRefresh(out error));
                TestAssert.False(host.CurrentViewModel.DvhSeries[0].IsSelected, "Same-plan refresh must preserve explicit hiding.");
                TestAssert.True(snapshot.DvhSeries[0].Selected, "Presentation state must not change the captured clinical facts.");
                host.CurrentViewModel.ResetDvhSelections();
                TestAssert.True(host.CurrentViewModel.DvhSeries[0].IsSelected);
            }
        }

        public static void HideUnmatchedLeavesMappingAndUnavailableMatches()
        {
            var snapshot = SyntheticScenarioFactory.Create("baseline-pass");
            snapshot.PqmRows[0].ResolvedStructureId = null;
            snapshot.PqmRows[1].ResolvedStructureId = "Matched organ";
            snapshot.PqmRows[1].Status = ReviewStatusCodes.NotEvaluated;
            dynamic workspace = new ReviewWorkspaceViewModel(snapshot);
            TestAssert.NotNull(((object)workspace).GetType().GetProperty("HideUnmatched"));
            int mappingCount = workspace.StructureMappings.Count;
            TestAssert.True((bool)workspace.HideUnmatched, "New workspaces must hide unmatched rows by default.");
            var shown = ((IEnumerable)workspace.VisiblePqmRows).Cast<ReviewPqmRowViewModel>().ToList();
            TestAssert.False(shown.Any(r => r.StableId == snapshot.PqmRows[0].StableId));
            TestAssert.True(shown.Any(r => r.StableId == snapshot.PqmRows[1].StableId));
            TestAssert.Equal(mappingCount, (int)workspace.StructureMappings.Count);
            TestAssert.Equal(snapshot.PqmRows.Count, (int)workspace.PqmRows.Count);
            workspace.HideUnmatched = false;
            TestAssert.Equal(snapshot.PqmRows.Count, ((IEnumerable)workspace.VisiblePqmRows).Cast<object>().Count(),
                "The user must still be able to show unmatched goals.");
            var report = new ClearPlan.Reporting.ReviewSnapshotReportMapper().Map(snapshot);
            TestAssert.True(new ClearPlan.Reporting.ReviewReportDocument().HideUnmatched && report.HideUnmatched,
                "Direct PDF/HTML exports must share the GUI default.");
            TestAssert.False(report.VisiblePqmRows().Any(r => r.StableId == snapshot.PqmRows[0].StableId));
            TestAssert.True(report.VisiblePqmRows().Any(r => r.StableId == snapshot.PqmRows[1].StableId));
            TestAssert.True(report.VisibilityDisclosure().Contains("unmatched"));
            report.HideUnmatched = false;
            TestAssert.Equal(snapshot.PqmRows.Count, report.VisiblePqmRows().Count);
        }

        public static void CompactChecksAndClickableLegendsAreWired()
        {
            string xaml = File.ReadAllText(Path.Combine("ClearPlan.Presentation", "Views", "ReviewWorkspaceView.xaml"));
            TestAssert.True(xaml.Contains("WorkspaceDvhLegend"), "A keyboard-accessible clickable right-hand legend is required.");
            TestAssert.True(xaml.Contains("HideUnmatched") && xaml.Contains("IncludeBeamEyeViews"));
            var document = System.Xml.Linq.XDocument.Parse(xaml);
            System.Xml.Linq.XNamespace ns = "http://schemas.microsoft.com/winfx/2006/xaml/presentation";
            foreach (var grid in document.Descendants(ns + "DataGrid").Where(g => (string)g.Attribute("ItemsSource") == "{Binding PlanCheckRows}"))
            {
                var columns = grid.Element(ns + "DataGrid.Columns").Elements().ToList();
                TestAssert.Equal("{lang:Translate Value='Meldung'}", (string)columns[0].Attribute("Header"));
                TestAssert.Equal("{lang:Translate Value='Status'}", (string)columns[1].Attribute("Header"));
                TestAssert.False(columns.Any(c => new[] { "Einheit", "Beobachtet", "Erwartet" }.Any(label => ((string)c.Attribute("Header") ?? "").Contains(label))));
            }
            var goal = typeof(ReviewPqmRowViewModel).GetProperty("GoalText");
            TestAssert.NotNull(goal, "Compact PQM must retain the strict comparator and unit.");
            var row = new ReviewPqmRowViewModel(new ReviewPqmRow { Comparator = ">", Goal = 57, Unit = "Gy" });
            TestAssert.Equal("> 57 Gy", (string)goal.GetValue(row));
            foreach (var grid in document.Descendants(ns + "DataGrid").Where(g => (string)g.Attribute("ItemsSource") == "{Binding VisiblePqmRows}"))
            {
                var columns = grid.Element(ns + "DataGrid.Columns").Elements().ToList();
                TestAssert.True(columns.Count <= 8, "PQM needs readable clinical values, not thirteen clipped columns.");
                TestAssert.True(columns.Any(c => (string)c.Attribute("Binding") == "{Binding GoalText}"));
            }
        }
    }
}
